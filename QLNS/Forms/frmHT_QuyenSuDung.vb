Public Class frmHT_QuyenSuDung

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _Systems As clsHeThong = New clsHeThong
    'Private _UserName As String = gUsername
    Private _IdCanBo As String = ""
    Private _IsEdit As Boolean = False 'Xác định đây là tạo user mới hay sửa
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler

    Public WriteOnly Property IdCanBo() As String
        Set(ByVal value As String)
            _IdCanBo = value
        End Set
    End Property

    Public WriteOnly Property Username() As String
        Set(ByVal value As String)
            txt_Username.Text = value
        End Set
    End Property

    Public WriteOnly Property MaPOS() As String
        Set(ByVal value As String)
            txtMaPOS.Text = value
        End Set
    End Property

    Public WriteOnly Property HoTen() As String
        Set(ByVal value As String)
            lbl_fullname.Text = value
        End Set
    End Property
#End Region

#Region "---> Events - Các sự kiện chính <---"
    Private Sub frmHT_DoiMkhau_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Fill dữ liệu combobox Nhóm quyền
        Dim conn As New DBAccess
        Dim StrSQL As String
        If Microsoft.VisualBasic.Left(Globals.Group, 1).ToUpper <> "U" Then
            StrSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Order by ma_nhom Asc")
        Else
            StrSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Where LEFT(ma_nhom,1) = 'U' Order by ma_nhom Asc")
        End If
        cbo_Binding(cboNhomQuyen, conn.getDataTable(StrSQL))
        If cboNhomQuyen.Items.Count > 0 Then cboNhomQuyen.SelectedIndex = 0

        If txt_Username.Text.Trim = "" Then
            lblTitle.Text = "TẠO MỚI THÔNG TIN ĐĂNG NHẬP"
            txt_Username.Text = _LayTenDangNhap()
            _IsEdit = False
            chkChangePass.Checked = True
            chkChangePass.Visible = False
        Else
            lblTitle.Text = "THAY ĐỔI QUYỀN SỬ DỤNG"
            _IsEdit = True
            chkChangePass.Checked = False
            'Lấy thông tin
            Dim dt As DataTable = conn.getDataTable("SELECT Login_POS, ID_Nhom FROM HS_CanBo WHERE Login_Username='" & txt_Username.Text.Trim & "'")
            If dt.Rows.Count > 0 Then
                txtMaPOS.Text = dt.Rows(0)("Login_POS")
                For i As Integer = cboNhomQuyen.Items.Count - 1 To 0 Step -1

                    cboNhomQuyen.SelectedIndex = i
                    If cboNhomQuyen.SelectedValue = dt.Rows(0)("ID_Nhom") Then Exit For
                Next
            End If

        End If
        chkChangePass_CheckedChanged(sender, e)
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Dispose()
    End Sub

    Private Sub frmHT_DoiMkhau_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Dispose()
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        Dim StrSQL As String
        If (IsValid()) Then
            'Lưu tài khoản đăng nhập với 3 thông tin: Username, Pass và mã POS để khi cán bộ đi đơn vị khác ko đăng nhập được
            If chkChangePass.Checked Then
                Dim _passWord As String = System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Globals.Find_Replace(edt_pass_new.Text.Trim().ToString()), "MD5")
                StrSQL = String.Format("UPDATE HS_CanBo SET login_Username = N'{0}', login_Password = N'{1}', login_POS = N'{2}', ID_Nhom = N'{3}' WHERE idCanBo = '" & _IdCanBo & "'", _
                                                     txt_Username.Text.Trim, _passWord, txtMaPOS.Text.Trim, cboNhomQuyen.SelectedValue)

            Else
                StrSQL = String.Format("UPDATE HS_CanBo SET login_Username = N'{0}', login_POS = N'{1}', ID_Nhom = N'{2}' WHERE idCanBo = '" & _IdCanBo & "'", _
                                                     txt_Username.Text.Trim, txtMaPOS.Text.Trim, cboNhomQuyen.SelectedValue)
            End If
            Dim conn As New DBAccess
            conn.executeSQL(StrSQL)
            MessageBox.Show("Thiết lập tài khoản đăng nhập thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            edt_pass_new.Text = ""
            edt_confirm.Text = ""
        End If
    End Sub

    Private Sub chkChangePass_CheckedChanged(sender As Object, e As EventArgs) Handles chkChangePass.CheckedChanged
        edt_pass_new.Enabled = chkChangePass.Checked
        edt_confirm.Enabled = chkChangePass.Checked
    End Sub
#End Region

#Region "Local functions"
    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra hợp lệ dữ liệu cập nhật
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Dim conn As New DBAccess

        If Not _IsEdit AndAlso conn.getNumber("SELECT Count(*) FROM HS_CanBo WHERE Login_Username = '" & txt_Username.Text & "'") > 0 Then
            MessageBox.Show("Tên đăng nhập """ & txt_Username.Text & """ đã có, thử lại với tên đăng nhập mới!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            txt_Username.Text = _LayTenDangNhap()
            ActiveControl = edt_pass_new
            Return False
        ElseIf chkChangePass.Checked AndAlso (edt_pass_new.Text.Trim() = "") Then
            MessageBox.Show("Mật khẩu mới của thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_pass_new
            Return False
        ElseIf chkChangePass.Checked AndAlso (edt_pass_new.Text.Trim().Length > 64 Or edt_pass_new.Text.Trim().Length < 5) Then
            If (edt_pass_new.Text.Trim().Length > 64) Then
                MessageBox.Show("Mật khẩu thành viên không được vượt quá 64 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_pass_new
                Return False
            End If
            If (edt_pass_new.Text.Trim().Length < 5) Then
                MessageBox.Show("Mật khẩu thành viên không được nhỏ hơn 5 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_pass_new
                Return False
            End If
        ElseIf chkChangePass.Checked AndAlso (edt_confirm.Text.Trim() = "") Then
            MessageBox.Show("Xác nhận lại mật khẩu mới của thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_confirm
            Return False
        ElseIf chkChangePass.Checked AndAlso (edt_pass_new.Text.Trim() <> edt_confirm.Text.Trim()) Then
            MessageBox.Show("Xác nhận mật khẩu phải trùng với mật khẩu vừa mới thiết lập!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_confirm
            Return False
        ElseIf (cboNhomQuyen.SelectedIndex < 0) Then
            MessageBox.Show("Chưa chọn nhóm quyền!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_confirm
            Return False
        ElseIf cboNhomQuyen.SelectedValue.ToString.ToUpper = "OPERATOR" AndAlso DONVI <> gMaDonViTW Then
            'Nếu không phải cán bộ Admin ở HSC thì không được thiết lập quyền Admin cho cấp khác
            My_MessageBox("Thiết lập quyền quản trị không dành cho đơn vị cấp Chi nhánh!")
            ActiveControl = edt_confirm
            Return False
        End If

        Return True
    End Function

    Function _LayTenDangNhap() As String
        Dim conn As New DBAccess
        Return conn.getString("SELECT dbo.funcFullName2Username(dbo.funcSigned2Unsigned(N'" & lbl_fullname.Text & "'))")
    End Function
#End Region

End Class