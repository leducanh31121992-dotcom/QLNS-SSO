Module modVar

    ''' <summary>
    ''' Hàm lấy giá trị khởi tạo cho các biến dùng chung
    ''' </summary>
    ''' <remarks></remarks>

    Public CAP As String = ""
    Public Cap_Nd As Integer = 0
    Public DONVI As String = ""
    Public TEN_DV_VT As String = ""
    Public BrandNameByUserLogin As String = ""
    Public IdDONVI As Integer = 0
    Public TRUCTHUOC As Byte
    Public DIABAN As String = ""
    Public ALL As Boolean
    Public MAX_KY As Byte = 1
    Public GIAMDOC As String = ""
    Public PHOGIAMDOC As String = ""
    Public KETOANTRUONG As String = ""
    Public TRUONGHCTC As String = ""
    Public NGUOILAPBIEU As String = ""

    Public gID_CanBo As String = ""
    Public gUsername As String = ""

    'Biến tạm, 
    Public KT_ChuyenMon As Int16 = 1
    'Khen thuong cua Chuyen mon: KT_ChuyenMon=1
    'Khen thuong cua Cong doan: KT_ChuyenMon=0

    'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
    Public g_oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture

    Public Sub initPublicVar()
        Try
            Dim dbconn As DBAccess = New DBAccess

            Dim iRootId As Integer = SoftSqlHelper.GetNumber(String.Format("Select Id_Goc From ChiNhanh Where Ma_So = '{0}'", DONVI), 0)
            Cap_Nd = IIf(iRootId = 0 Or DONVI = "000199", 1, IIf(iRootId = 1, 2, 3))
            BrandNameByUserLogin = SoftSqlHelper.GetString(String.Format("Select Top 1 Ten_Goi From ChiNhanh Where Ma_So = '{0}' Order By Id", DONVI), "")
            CAP = IIf(DONVI = gMaDonViTW, 1, 2)
            'DONVI = getSystemVar("DONVI")
            TEN_DV_VT = dbconn.getString("SELECT ten_vt FROM ChiNhanh WHERE ma_so = '" & DONVI & "'") ' getSystemVar("TEN_VT")
            DIABAN = My_CStr(IIf(DONVI = gMaDonViTW, "Hà Nội", dbconn.getString("SELECT T2.ten_goi FROM ChiNhanh T1 INNER JOIN DiaDanh T2 ON T1.ma_so = T2.ma_so WHERE T1.ma_so = '" & DONVI & "'")), "") 'getSystemVar("DIABAN")

            ALL = 1 'CBool(getSystemVar("ALL"))
            IdDONVI = dbconn.getNumber("SELECT id FROM ChiNhanh WHERE ma_so='" & DONVI.Trim & "'")
            TRUCTHUOC = IIf(DONVI = gMaDonViTW, 1, 0) 'CByte(getSystemVar("TRUCTHUOC"))

            Dim _SQLHelper As New DBAccess
            Dim StrSQL As String = "SELECT Top 1 * FROM SysVar WHERE ID_DonVi = " & IdDONVI & " Order By NgayHL Desc,TrangThai Desc"
            Dim dt As DataTable = _SQLHelper.getDataTable(StrSQL)
            If dt.Rows.Count > 0 Then
                MAX_KY = Math.Min(My_CInt(dt.Rows(0)("MAX_KY"), 1), 2)
                'THUETNCN = Math.Min(My_CInt(dt.Rows(0)("TINH_THUE_TNCN"), 1), 1)
                GIAMDOC = dt.Rows(0)("GIAMDOC")
                PHOGIAMDOC = dt.Rows(0)("PHOGIAMDOC")
                TRUONGHCTC = dt.Rows(0)("TRUONGHCTC")
                KETOANTRUONG = dt.Rows(0)("KETOANTRUONG")
                DIABAN = My_CStr(dt.Rows(0)("DIABAN"), "")
                NGUOILAPBIEU = MainForm.lblNguoiSuDung.Text
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

End Module

