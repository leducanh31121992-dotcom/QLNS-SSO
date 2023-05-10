Public Class frmHT_DoiMkhau

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _Systems As clsHeThong = New clsHeThong
    Private _UserName As String = gUsername

    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
#End Region

#Region "---> Events - Các sự kiện chính <---"
    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra hợp lệ dữ liệu cập nhật
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        If (edt_pass_old.Text.Trim() = "") Then
            MessageBox.Show("Mật khẩu cũ của thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_pass_old
            Return False
        End If
        Dim dbconn As New DBAccess
        If dbconn.getNumber("SELECT COUNT(*) FROM HS_CanBo WHERE idDonVi = " & IdDONVI & " AND login_Username = '" & gUsername & "' AND login_password = '" & System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(edt_pass_old.Text.Trim, "MD5") & "'") = 0 Then
            MessageBox.Show("Mật khẩu cũ của không đúng, vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_pass_old
            Return False
        End If
        If (edt_pass_new.Text.Trim() = "") Then
            MessageBox.Show("Mật khẩu mới của thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_pass_new
            Return False
        End If
        If (edt_pass_new.Text.Trim().Length > 64 Or edt_pass_new.Text.Trim().Length < 5) Then
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
        End If
        If (edt_confirm.Text.Trim() = "") Then
            MessageBox.Show("Xác nhận lại mật khẩu mới của thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_confirm
            Return False
        End If
        If (edt_pass_new.Text.Trim() <> edt_confirm.Text.Trim()) Then
            MessageBox.Show("Xác nhận mật khẩu phải trùng với mật khẩu vừa mới thiết lập!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_confirm
            Return False
        End If
        Return True
    End Function

    Private Sub frmHT_DoiMkhau_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lbl_fullname.Text = MainForm.lblNguoiSuDung.Text
        lbl_username.Text = gUsername
        edt_pass_old.Text = ""
        edt_pass_new.Text = ""
        edt_confirm.Text = ""
        ActiveControl = edt_pass_old
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Dispose()
    End Sub

    Private Sub frmHT_DoiMkhau_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Dispose()
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        If (IsValid()) Then
            Dim _passWord As String = System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Globals.Find_Replace(edt_pass_new.Text.Trim().ToString()), "MD5")
            _Systems.Update_Password(lbl_username.Text.Trim(), _passWord)
            MessageBox.Show("Bạn đã sửa đổi mật khẩu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            edt_pass_old.Text = ""
            edt_pass_new.Text = ""
            edt_confirm.Text = ""
        End If
    End Sub

    Private Sub edt_pass_old_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_pass_old.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_pass_old.Text.Trim() <> "") Then
                edt_pass_new.Focus()
            Else
                edt_pass_old.Focus()
            End If
        End If
    End Sub

    Private Sub edt_pass_new_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_pass_new.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_pass_new.Text.Trim() <> "") Then
                edt_confirm.Focus()
            Else
                edt_pass_new.Focus()
            End If
        End If
    End Sub

    Private Sub edt_pass_new_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_pass_new.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_pass_old.Focus()
        End If
    End Sub

    Private Sub edt_confirm_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_confirm.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_confirm.Text.Trim() <> "") Then
                btn_accept.Focus()
            Else
                edt_confirm.Focus()
            End If
        End If
    End Sub

    Private Sub edt_confirm_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_confirm.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_pass_new.Focus()
        End If
    End Sub
#End Region

End Class