Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Public Class KHLD_MangLuoiBLL
    ''' <summary>
    ''' Hàm thiết lập định dạng hiển thị của Cột dữ liệu trong lưới DataGridView
    ''' </summary>
    ''' <param name="pNameDGV">Tên DataGridView</param>
    ''' <param name="pListColumns">Danh sách tên các cột của DataGridView</param>
    ''' <param name="pDelimiter">Ký tự phân cách tên các cột</param>
    ''' <param name="pAlignment">Chỉ số xác định Style Alignment của Cột. 1 - Trái; 2 - Giữa; 3 - Phải</param>
    ''' <remarks></remarks>
    Private Sub FormatStyleAlignment_ColumnDataGridView(ByVal pNameDGV As DataGridView, ByVal pListColumns As String, ByVal pDelimiter As String, pAlignment As Byte)
        Dim ARL_ColRight() As String = Globals.Splip_Strings(pListColumns, pDelimiter)
        For Each _ValueNull As String In ARL_ColRight
            If Not String.IsNullOrEmpty(_ValueNull) Then
                If pAlignment = 1 Then
                    pNameDGV.Columns("cln_" + _ValueNull).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                ElseIf pAlignment = 2 Then
                    pNameDGV.Columns("cln_" + _ValueNull).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ElseIf pAlignment = 3 Then
                    pNameDGV.Columns("cln_" + _ValueNull).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Hàm định nghĩa lưới dữ liệu - Liên quan đến quản lý KHLĐ - Mạng lưới đơn vị
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="_GridIndex">
    '''               0: Lưới dữ liệu Mạng lưới đơn vị
    '''               1: Kế hoạch lao động
    ''' </param>
    ''' <remarks></remarks>
    ''' 
    Public Sub Create_Frame_GridView(ByVal dgv_name As DataGridView, ByVal _GridIndex As Byte)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        Dim sColumnRight As String = ""
        Select Case _GridIndex
            Case 0       'Định nghĩa lưới dữ liệu - Lưới dữ liệu Mạng lưới đơn vị
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_IdKHML", "Mã hiệu")
                dgv_name.Columns.Add("cln_Id", "Id")
                dgv_name.Columns.Add("cln_KieuIn", "KieuIn")

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_ThoiDiem", "Ngày thông báo")
                dgv_name.Columns.Add("cln_DonVi", "Đơn vị")
                dgv_name.Columns.Add("cln_SoXaPhuong", "Số xã phường")
                dgv_name.Columns.Add("cln_SoDiemGD", "Số điểm giao dịch")
                dgv_name.Columns.Add("cln_SoToTKVV", "Số Tổ TK&VV")
                dgv_name.Columns.Add("cln_SoKH_DN", "Số hộ còn dư nợ")
                dgv_name.Columns.Add("cln_TongDN", "Tổng dư nợ") '11
                dgv_name.Columns.Add("cln_DuNo_TH", "Trong hạn") '12
                dgv_name.Columns.Add("cln_DuNo_QH", "Quá hạn")
                dgv_name.Columns.Add("cln_DuNo_KH", "Khoanh") '14
                dgv_name.Columns.Add("cln_GhiChu", "Ghi chú")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_ThoiDiem").Width = 85
                dgv_name.Columns("cln_DonVi").Width = 220
                dgv_name.Columns("cln_SoXaPhuong").Width = 65
                dgv_name.Columns("cln_SoDiemGD").Width = 65
                dgv_name.Columns("cln_SoToTKVV").Width = 70
                dgv_name.Columns("cln_SoKH_DN").Width = 80
                dgv_name.Columns("cln_TongDN").Width = 105
                dgv_name.Columns("cln_DuNo_TH").Width = 95
                dgv_name.Columns("cln_DuNo_QH").Width = 78
                dgv_name.Columns("cln_DuNo_KH").Width = 70
                dgv_name.Columns("cln_GhiChu").Width = 180

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 15
                    dgv_name.Columns(i).ReadOnly = True
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    If i >= 11 And i <= 14 Then
                        dgv_name.Columns(i).HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
                    Else
                        dgv_name.Columns(i).HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                    End If
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_ThoiDiem").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                sColumnRight = "SoXaPhuong;SoDiemGD;SoToTKVV;SoKH_DN;TongDN;DuNo_TH;DuNo_QH;DuNo_KH"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)

                dgv_name.ColumnHeadersHeight = 45
                dgv_name.Columns("cln_Choice").Visible = False

            Case 1      'Định nghĩa lưới dữ liệu - Kế hoạch lao động
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_IdKHML", "Mã hiệu")
                dgv_name.Columns.Add("cln_Id", "Id")
                dgv_name.Columns.Add("cln_KieuIn", "KieuIn")
                dgv_name.Columns.Add("cln_IdDonVi", "IdDonVi")
                dgv_name.Columns.Add("cln_ChiNhanh_Id", "ChiNhanh_Id")
                dgv_name.Columns.Add("cln_IdPhongBan", "IdPhongBan")

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_ThoiDiem", "Thời điểm")
                dgv_name.Columns.Add("cln_Name", "Đơn vị/Phòng ban")
                dgv_name.Columns.Add("cln_SoLD_DaiHan", "Số lao động dài hạn")
                dgv_name.Columns.Add("cln_SoLD_NganHan", "Số lao động ngắn hạn")
                dgv_name.Columns.Add("cln_SoCV_ThBao", "Số công văn TB")
                dgv_name.Columns.Add("cln_GhiChu", "Ghi chú")
                dgv_name.Columns.Add("cln_Loai_DL", "Loai_DL")
                dgv_name.Columns.Add("cln_DonVi_Cd", "DonVi_Cd")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_ThoiDiem").Width = 85
                dgv_name.Columns("cln_Name").Width = 210
                dgv_name.Columns("cln_SoLD_DaiHan").Width = 90
                dgv_name.Columns("cln_SoLD_NganHan").Width = 80
                dgv_name.Columns("cln_SoCV_ThBao").Width = 180
                dgv_name.Columns("cln_GhiChu").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 7 To 15
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_ThoiDiem").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoLD_DaiHan").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_SoLD_NganHan").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight

                dgv_name.ColumnHeadersHeight = 37
                dgv_name.Columns("cln_Choice").Visible = False
                dgv_name.Columns("cln_IdDonVi").Visible = False
                dgv_name.Columns("cln_ChiNhanh_Id").Visible = False
                dgv_name.Columns("cln_IdPhongBan").Visible = False
                dgv_name.Columns("cln_Loai_DL").Visible = False
                dgv_name.Columns("cln_DonVi_Cd").Visible = False
            Case 2       'Định nghĩa lưới dữ liệu - Nhu cầu lao động cần bổ sung và Số cán bộ được thông báo trúng tuyển vào chi nhánh trong tháng đang tham gia lớp đào tạo
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_IdKHML", "Mã hiệu")                               '1
                dgv_name.Columns.Add("cln_Id", "Id")                                        '2
                dgv_name.Columns.Add("cln_KieuIn", "KieuIn")                                '3

                dgv_name.Columns.Add("cln_STT", "STT")                                      '4
                dgv_name.Columns.Add("cln_ThoiDiem", "Thời điểm")                           '5
                dgv_name.Columns.Add("cln_DonVi_HT", "Đơn vị")                                 '6
                dgv_name.Columns.Add("cln_SoLD_ThgTruoc", "Lao động CMNV thực hiện tháng trước")        '7
                dgv_name.Columns.Add("cln_SoLD_ThgBC_DaiHan", "Lao động CMNV thực hiện trong tháng")    '8
                dgv_name.Columns.Add("cln_SoLD_ThgBC_DangLV", "Lao động đang làm việc" + vbNewLine + vbNewLine)                 '9
                dgv_name.Columns.Add("cln_SoLD_ThgBC_Nghi_HuongBHXH", vbNewLine + "Lao động nghỉ ốm đau, TS trong chế độ (được hưởng BHXH)")        '10
                dgv_name.Columns.Add("cln_SoLD_ThgBC_Nghi_KhongBHXH", vbNewLine + "Lao động nghỉ tự túc, Không lương, ốm đau, TS không trong chế độ (Không hưởng BHXH)")        '11
                dgv_name.Columns.Add("cln_SoLD_ThgBC_NganHan", "Lao động ngắn hạn thực hiện đến cuối tháng")                                                        '12 
                dgv_name.Columns.Add("cln_SoLD_ThgBC_NganHan_DB", "Định biên" + vbNewLine + vbNewLine + vbNewLine)                                                                                      '13
                dgv_name.Columns.Add("cln_SoLD_ThgBC_NganHan_PT", "Phụ trợ" + vbNewLine + vbNewLine + vbNewLine)                                                                                        '14
                dgv_name.Columns.Add("cln_TongDN", "Tổng dư nợ" + vbNewLine + "(tỷ đồng)" + vbNewLine + vbNewLine)                                                                          '15
                dgv_name.Columns.Add("cln_SoKH_DN", "Số khách hàng" + vbNewLine + vbNewLine)                                                                                                '16
                dgv_name.Columns.Add("cln_SoXaPhuong", "Số xã, phường, thị trấn" + vbNewLine + vbNewLine)                                                                                   '17
                dgv_name.Columns.Add("cln_SoDiemGD", "Số điểm giao dịch xã" + vbNewLine + vbNewLine)                                                                                        '18
                dgv_name.Columns.Add("cln_SoToTKVV", "Số Tổ TK&VV" + vbNewLine + vbNewLine)                                                                                                 '19
                dgv_name.Columns.Add("cln_SoLD_GiamTuNhien_KhacCN", "Số LĐ giảm tự nhiên trong tháng (như: nghỉ hưu, chuyển công tác,...)")                                '20
                dgv_name.Columns.Add("cln_ThBaoTTDangDT_SL", "Số lượng" + vbNewLine)                     'Số cán bộ được thông báo trúng tuyển vào chi nhánh trong tháng đang tham gia lớp đào tạo      '21
                dgv_name.Columns.Add("cln_ThBaoTTDangDT_SoCV", "Văn bản thông báo")                                                                                 '22
                dgv_name.Columns.Add("cln_NhuCauBS_SL", "Số lượng" + vbNewLine)                         'Nhu cầu lao động cần bổ sung                                           '23
                dgv_name.Columns.Add("cln_NhuCauBS_SoCV", "Văn bản thông báo")                                                                                      '24
                dgv_name.Columns.Add("cln_GhiChu", "Ghi chú")                                                                                                       '25

                sColumnRight = "Loai_DL;Loai_DL_HT;IdDonVi;DonVi_Cd;ChiNhanh_Id;ChiNhanh_HT;IdPhongBan;PhongBan_HT;DuNo_TH;DuNo_QH;DuNo_KH;SoLD_DaiHan;SoLD_NganHan;SoCV_ThBao;SoLD_GiamTuNhien;SoLD_GiamTN_01;SoLD_GiamTN_KhacCN_01;SoLD_GiamTN_02;SoLD_GiamTN_KhacCN_02"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnRight, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ThoiDiem").Width = 85
                dgv_name.Columns("cln_ThoiDiem").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DonVi_HT").Width = 160
                dgv_name.Columns("cln_DonVi_HT").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoLD_ThgTruoc").Width = 80
                dgv_name.Columns("cln_SoLD_ThgTruoc").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoLD_ThgBC_DaiHan").Width = 80
                dgv_name.Columns("cln_SoLD_ThgBC_DaiHan").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoLD_ThgBC_DangLV").Width = 80
                dgv_name.Columns("cln_SoLD_ThgBC_Nghi_HuongBHXH").Width = 80
                dgv_name.Columns("cln_SoLD_ThgBC_Nghi_KhongBHXH").Width = 80
                dgv_name.Columns("cln_SoLD_ThgBC_NganHan").Width = 80
                dgv_name.Columns("cln_SoLD_ThgBC_NganHan").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoLD_ThgBC_NganHan_DB").Width = 60
                dgv_name.Columns("cln_SoLD_ThgBC_NganHan_PT").Width = 60
                dgv_name.Columns("cln_TongDN").Width = 80
                dgv_name.Columns("cln_SoKH_DN").Width = 75
                dgv_name.Columns("cln_SoXaPhuong").Width = 60
                dgv_name.Columns("cln_SoDiemGD").Width = 60
                dgv_name.Columns("cln_SoToTKVV").Width = 70
                dgv_name.Columns("cln_SoLD_GiamTuNhien_KhacCN").Width = 70
                dgv_name.Columns("cln_SoLD_GiamTuNhien_KhacCN").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ThBaoTTDangDT_SL").Width = 65
                dgv_name.Columns("cln_ThBaoTTDangDT_SoCV").Width = 110
                dgv_name.Columns("cln_NhuCauBS_SL").Width = 65
                dgv_name.Columns("cln_NhuCauBS_SoCV").Width = 110
                dgv_name.Columns("cln_GhiChu").Width = 180
                dgv_name.Columns("cln_GhiChu").HeaderCell.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_ThoiDiem").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                sColumnRight = "SoLD_ThgTruoc;SoLD_ThgBC_DaiHan;SoLD_ThgBC_DangLV;SoLD_ThgBC_Nghi_HuongBHXH;SoLD_ThgBC_Nghi_KhongBHXH;SoLD_ThgBC_NganHan;SoLD_ThgBC_NganHan_DB;SoLD_ThgBC_NganHan_PT;TongDN;SoKH_DN;SoXaPhuong;SoDiemGD;SoToTKVV;SoLD_GiamTuNhien;ThBaoTTDangDT_SL;ThBaoTTDangDT_SoCV;NhuCauBS_SL;NhuCauBS_SoCV"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)
                'Để column check có thể edit
                sColumnRight = "Choice;ThBaoTTDangDT_SL;ThBaoTTDangDT_SoCV;NhuCauBS_SL;NhuCauBS_SoCV;GhiChu"
                ARL_Cols = Globals.Splip_Strings(sColumnRight, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).ReadOnly = False
                        dgv_name.Columns("cln_" + _Value).DefaultCellStyle.BackColor = Color.White
                    End If
                Next
                dgv_name.Columns("cln_Choice").Visible = False
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_ThoiDiem").Frozen = True
                dgv_name.Columns("cln_DonVi_HT").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_ThoiDiem").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_DonVi_HT").DefaultCellStyle.BackColor = Color.LightGray
        End Select
        dgv_name.Columns("cln_IdKHML").Visible = False
        dgv_name.Columns("cln_Id").Visible = False
        dgv_name.Columns("cln_KieuIn").Visible = False
    End Sub

    Public Function GetKhLd_MangLuoi_Search(ByVal _FlagTWCN As Byte, ByVal _ThoiDiem As String, ByVal _Id As Integer, ByVal _IdKHML As String, ByVal _Loai_DL As Byte, ByVal _ChiNhanh_Id As Integer, ByVal _DonVi_Cd As String) As DataTable
        Try
            Dim ds_ret As DataSet = New DataSet()
            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()

                Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("GetKHLD_MangLuoi_Search", conn_obj)
                    _command.CommandType = CommandType.StoredProcedure

                    _command.Parameters.Add(New SqlParameter("@pFlagTWCN", SqlDbType.TinyInt))
                    _command.Parameters("@pFlagTWCN").Value = _FlagTWCN
                    _command.Parameters.Add(New SqlParameter("@pThoiDiem", SqlDbType.VarChar, 10))
                    _command.Parameters("@pThoiDiem").Value = _ThoiDiem
                    _command.Parameters.Add(New SqlParameter("@pId", SqlDbType.Int))
                    _command.Parameters("@pId").Value = _Id
                    _command.Parameters.Add(New SqlParameter("@pIdKHML", SqlDbType.VarChar, 32))
                    _command.Parameters("@pIdKHML").Value = _IdKHML
                    _command.Parameters.Add(New SqlParameter("@pLoai_DL", SqlDbType.TinyInt))
                    _command.Parameters("@pLoai_DL").Value = _Loai_DL
                    _command.Parameters.Add(New SqlParameter("@pChiNhanh_Id", SqlDbType.Int))
                    _command.Parameters("@pChiNhanh_Id").Value = _ChiNhanh_Id
                    _command.Parameters.Add(New SqlParameter("@pDonVi_Cd", SqlDbType.VarChar, 32))
                    _command.Parameters("@pDonVi_Cd").Value = _DonVi_Cd

                    Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                        _sqldap.Fill(ds_ret, "KHLD_MangLuoi_TMP")
                        If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                            Return Nothing
                        End If
                    End Using
                End Using

            End Using
            Return ds_ret.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Lỗi lấy danh sách bản ghi KHLĐ-Mạng lưới: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    Public Function Insert_Update_KhLd_MangLuoi(ByVal _KHLD_MangLuoi As KHLD_MangLuoi.KhLd_MangLuoi) As String
        Dim _Ret As String = ""
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "KHLD_MangLuoi_Insert_Update"

                    _command.Parameters.Add("@_IdOutPut", SqlDbType.VarChar, 32).Direction = ParameterDirection.Output

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdKHML", _KHLD_MangLuoi.IdKHML, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThoiDiem", _KHLD_MangLuoi.ThoiDiem, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Loai_DL", _KHLD_MangLuoi.Loai_DL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdDonVi", _KHLD_MangLuoi.IdDonVi, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonVi_Cd", _KHLD_MangLuoi.DonVi_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdPhongBan", _KHLD_MangLuoi.IdPhongBan, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoLD_DaiHan", _KHLD_MangLuoi.SoLD_DaiHan, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoLD_NganHan", _KHLD_MangLuoi.SoLD_NganHan, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_So_XaPhuong", _KHLD_MangLuoi.So_XaPhuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_So_DiemGD", _KHLD_MangLuoi.So_DiemGD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_So_ToTKVV", _KHLD_MangLuoi.So_ToTKVV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoKH_DN", _KHLD_MangLuoi.SoKH_DN, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TongDN", _KHLD_MangLuoi.TongDN, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DuNo_TH", _KHLD_MangLuoi.DuNo_TH, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DuNo_QH", _KHLD_MangLuoi.DuNo_QH, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DuNo_KH", _KHLD_MangLuoi.DuNo_KH, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoCV_ThBao", _KHLD_MangLuoi.SoCV_ThBao, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NhuCauBS_SL", _KHLD_MangLuoi.NhuCauBS_SL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NhuCauBS_SoCV", _KHLD_MangLuoi.NhuCauBS_SoCV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThBaoTTDangDT_SL", _KHLD_MangLuoi.ThBaoTTDangDT_SL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThBaoTTDangDT_SoCV", _KHLD_MangLuoi.ThBaoTTDangDT_SoCV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_GhiChu", _KHLD_MangLuoi.GhiChu, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_CreatedBy", _KHLD_MangLuoi.CreatedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _KHLD_MangLuoi.ModifiedBy, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _KHLD_MangLuoi.IdOutPut = _command.Parameters("@_IdOutPut").Value.ToString()
                        _Ret = _KHLD_MangLuoi.IdOutPut
                    End If

                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Cập nhật dữ liệu KHLĐ, Mạng lưới đơn vị, Nhu cầu tuyển dụng (Insert_Update_KhLd_MangLuoi)", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = 0
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    Public Function Delete_KhLd_MangLuoi(ByVal _IdKHML As String, ByVal _UserName As String) As Boolean
        Dim _Ret As Boolean = False

        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "KHLD_MangLuoi_Delete"

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdKHML", _IdKHML, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_UserName", _UserName, ParameterDirection.Input))
                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _Ret = True
                    End If
                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Xóa Kế hoạch LĐ - Mạng lưới đơn vị", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = False
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    Public Function GetListThoiDiem(ByVal pLoai_DL As Byte, ByVal pNone As String, ByVal pPosRun As String) As DataTable
        Dim sSQL As String = ""
        If (String.IsNullOrEmpty(pNone)) Then
            If pLoai_DL <> 3 Then
                sSQL = ""
                sSQL = sSQL & "Select Distinct Row_Number() Over(Partition By Loai_DL Order By ThoiDiem Desc) Value,Convert(Varchar(10), KK.ThoiDiem, 105) Display From"
                sSQL = sSQL & " 	   ("
                sSQL = sSQL & " 	        Select Distinct Loai_DL,ThoiDiem From KHLD_MangLuoi"
                sSQL = sSQL & " 	   ) As KK Where KK.Loai_DL = " & pLoai_DL
            Else
                sSQL = ""
                sSQL = sSQL & "Select Distinct Row_Number() Over(Order By ThoiDiem Desc) Value,Convert(Varchar(10), KK.ThoiDiem, 105) Display From"
                sSQL = sSQL & " 	   ("
                sSQL = sSQL & " 	        Select Distinct Loai_DL,ThoiDiem From KHLD_MangLuoi Where DonVi_Cd In (Select Ma_So From ChiNhanh Where Status=1 And (Id In (Select Id From ChiNhanh Where Ma_So = '" & pPosRun & "') Or Id_Goc In (Select Id From ChiNhanh Where Ma_So = '" & pPosRun & "')))"
                sSQL = sSQL & " 	   ) As KK Where KK.Loai_DL <> 1"
            End If
        Else
            If pLoai_DL <> 3 Then
                sSQL = "Select 0 As Value, N'" & pNone & "' As Display Union "
                sSQL = sSQL & "Select Distinct Row_Number() Over(Partition By Loai_DL Order By ThoiDiem Desc) Value,Convert(Varchar(10), KK.ThoiDiem, 105) Display From"
                sSQL = sSQL & " 	   ("
                sSQL = sSQL & " 	        Select Distinct Loai_DL,ThoiDiem From KHLD_MangLuoi "
                If pPosRun <> "000100" And pPosRun <> "000199" Then
                    sSQL = sSQL & String.Format(" Where DonVi_Cd Like '{0}%'", pPosRun.Substring(0, 4))
                End If
                sSQL = sSQL & " 	   ) As KK Where KK.Loai_DL = " & pLoai_DL
            Else
                sSQL = "Select 0 As Value, N'" & pNone & "' As Display Union "
                sSQL = sSQL & "Select Distinct Row_Number() Over(Order By ThoiDiem Desc) Value,Convert(Varchar(10), KK.ThoiDiem, 105) Display From"
                sSQL = sSQL & " 	   ("
                sSQL = sSQL & " 	        Select Distinct Loai_DL,ThoiDiem From KHLD_MangLuoi Where DonVi_Cd In (Select Ma_So From ChiNhanh Where Status=1 And (Id In (Select Id From ChiNhanh Where Ma_So = '" & pPosRun & "') Or Id_Goc In (Select Id From ChiNhanh Where Ma_So = '" & pPosRun & "')))"
                sSQL = sSQL & " 	   ) As KK Where KK.Loai_DL <> 1"
            End If
        End If
        Try
            Using db_List As DataTable = SoftSqlHelper.ExecuteForSQL(sSQL)
                Return db_List
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi lấy Danh sách thời điểm thống kê KHLĐ - Mạng lưới: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm lấy dữ liệu báo cáo: TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI 
    ''' </summary>
    ''' <param name="_FlagTWCN"></param>
    ''' <param name="_ThoiDiem">Thời điểm lấy dữ liệu -> Ngày cuối tháng (Bắt buộc phải truyền)</param>
    ''' <param name="_Id"></param>
    ''' <param name="_IdKHML"></param>
    ''' <param name="_Loai_DL">Loại dữ liệu là loại 4</param>
    ''' <param name="_ChiNhanh_Id">Nếu là cấp chi nhánh thì truyền hoặc truyền DonVi_Cd</param>
    ''' <param name="_DonVi_Cd">Nếu là cấp chi nhánh thì truyền hoặc truyền ChiNhanh_Id</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function KHLD_MangLuoi_GetSearchBC(ByVal _FlagTWCN As Byte, ByVal _ThoiDiem As String, ByVal _Id As Integer, ByVal _IdKHML As String, ByVal _Loai_DL As Byte, ByVal _ChiNhanh_Id As Integer, ByVal _DonVi_Cd As String) As DataTable
        Try
            Dim ds_ret As DataSet = New DataSet()
            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()

                Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("KHLD_MangLuoi_GetSearchBC", conn_obj)
                    _command.CommandType = CommandType.StoredProcedure

                    _command.Parameters.Add(New SqlParameter("@pFlagTWCN", SqlDbType.TinyInt))
                    _command.Parameters("@pFlagTWCN").Value = _FlagTWCN
                    _command.Parameters.Add(New SqlParameter("@pThoiDiem", SqlDbType.VarChar, 10))
                    _command.Parameters("@pThoiDiem").Value = _ThoiDiem
                    _command.Parameters.Add(New SqlParameter("@pId", SqlDbType.Int))
                    _command.Parameters("@pId").Value = _Id
                    _command.Parameters.Add(New SqlParameter("@pIdKHML", SqlDbType.VarChar, 32))
                    _command.Parameters("@pIdKHML").Value = _IdKHML
                    _command.Parameters.Add(New SqlParameter("@pLoai_DL", SqlDbType.TinyInt))
                    _command.Parameters("@pLoai_DL").Value = _Loai_DL
                    _command.Parameters.Add(New SqlParameter("@pChiNhanh_Id", SqlDbType.Int))
                    _command.Parameters("@pChiNhanh_Id").Value = _ChiNhanh_Id
                    _command.Parameters.Add(New SqlParameter("@pDonVi_Cd", SqlDbType.VarChar, 32))
                    _command.Parameters("@pDonVi_Cd").Value = _DonVi_Cd

                    Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                        _sqldap.Fill(ds_ret, "KHLD_MangLuoi_TMP")
                        If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                            Return Nothing
                        End If
                    End Using
                End Using

            End Using
            Return ds_ret.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Lỗi lấy danh sách bản ghi KHLĐ-Mạng lưới: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Hàm lấy dữ liệu hiển thị báo cáo Lao động - Mạng lưới (Mẫu số 09/BCLĐ)
    ''' </summary>
    ''' <param name="pCapBC">Cấp lấy số liệu. 1 - TW => Lấy đến Từng chi nhánh, từng PGD; 2 - Chi nhánh => Lấy Tổng chi nhánh và các PGD; 3 - PGD => Chỉ lấy PGD</param>
    ''' <param name="pThoiDiem">Thời điểm báo cáo: Bắt buộc phải truyền vào. Định dạng yyyy-MM-dd</param>
    ''' <param name="pDonViId">Đơn vị lấy SL. Nếu toàn quốc tổng hợp là 0</param>
    ''' <param name="pFlagCall">Cờ xác định đầu ra của dữ liệu BC: 1 - Báo cáo hằng tháng Tổng hợp TQ; 2 - Báo cáo hằng tháng chi tiết CN</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BaoCao_KHLDMangLuoi09(ByVal pCapBC As Byte, ByVal pThoiDiem As String, ByVal pDonViId As Integer, ByVal pFlagCall As Integer) As DataTable
        Try
            Dim ds_ret As DataSet = New DataSet()
            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()

                Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("BaoCao_KHLD_MangLuoi_09", conn_obj)
                    _command.CommandType = CommandType.StoredProcedure

                    _command.Parameters.Add(New SqlParameter("@pCapBC", SqlDbType.TinyInt))
                    _command.Parameters("@pCapBC").Value = pCapBC

                    _command.Parameters.Add(New SqlParameter("@pThoiDiem", SqlDbType.VarChar, 10))
                    _command.Parameters("@pThoiDiem").Value = pThoiDiem

                    _command.Parameters.Add(New SqlParameter("@pDonViId", SqlDbType.Int))
                    _command.Parameters("@pDonViId").Value = pDonViId

                    _command.Parameters.Add(New SqlParameter("@pFlagCall", SqlDbType.Int))
                    _command.Parameters("@pFlagCall").Value = pFlagCall

                    Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                        _sqldap.Fill(ds_ret, "KHLD_MangLuoi_TMP")
                        If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                            Return Nothing
                        End If
                    End Using
                End Using

            End Using
            Return ds_ret.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Lỗi lấy Báo cáo Lao động - Mạng lưới (Mẫu số 09/BCLĐ): " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function
End Class
