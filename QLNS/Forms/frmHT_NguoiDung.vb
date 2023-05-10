
Public Class frmHT_NguoiDung

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _Systems As clsHeThong = New clsHeThong
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler

    Private obj_grmembers As frmHT_NhomThVien

    Public userName As String
    Private arr_Groups As ArrayList = New ArrayList()
    Public FlagEvent As Byte

    Private _Node As String
    Public Property Node() As String
        Get
            Return _Node
        End Get
        Set(ByVal value As String)
            _Node = value
        End Set
    End Property

    Private strSQL As String = ""
#End Region

#Region "---> Functions: Các hàm dùng chung <---"
    Private Sub ResetAll_Controls()
        If (FlagEvent = 1) Then
            cb_mgroup.SelectedIndex = 0
            edt_fullname.Text = ""
            edt_username.Text = ""
            edt_password.Text = ""
            edt_email.Text = ""
            edt_note.Text = ""
        ElseIf (FlagEvent = 2) Then
            Dim dr As DataRow
            dr = _Systems.GetMember(userName)
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    cb_mgroup.SelectedIndex = IIf(dr("id_nhom").ToString() <> "", CType(arr_Groups.IndexOf(dr("id_nhom").ToString()), Int32), 0)
                    edt_fullname.Text = dr("ho_ten").ToString().Trim()
                    edt_username.Text = dr("ten_tv").ToString().Trim()
                    edt_password.Text = ""
                    edt_email.Text = dr("email").ToString().Trim()
                    edt_note.Text = dr("ghi_chu").ToString().Trim()
                End If
            End If
        End If
        ActiveControl = cb_mgroup
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra dữ liệu hợp lệ trước khi cập nhật
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        If (cb_mgroup.SelectedIndex <= 0 And cb_mgroup.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nhóm thành viên!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_mgroup
            Return False
        End If
        If (edt_fullname.Text.Trim() = "") Then
            MessageBox.Show("Họ và tên thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_fullname
            Return False
        End If
        If (edt_username.Text.Trim() = "") Then
            MessageBox.Show("Tên đăng nhập của thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_username
            Return False
        End If
        If (edt_password.Text.Trim() = "") Then
            MessageBox.Show("Mật khẩu của thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_password
            Return False
        End If
        'Bắt điều kiện quy định độ dài của tên đăng nhập và mật khẩu: Tên đăng nhập không dài quá 16 ký tự. Mật khẩu không nhỏ hơn 5 và không lớn quá 64 ký tự
        If (edt_username.Text.Trim().Length > 16) Then
            MessageBox.Show("Tên đăng nhập của thành viên không được dài quá 16 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_username
            Return False
        End If
        If (edt_username.Text.Trim().Length < 3) Then
            MessageBox.Show("Tên đăng nhập của thành viên không được nhỏ hơn 3 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_username
            Return False
        End If
        If (edt_password.Text.Trim().Length > 64 Or edt_password.Text.Trim().Length < 5) Then
            If (edt_password.Text.Trim().Length > 64) Then
                MessageBox.Show("Mật khẩu thành viên không được vượt quá 64 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_password
                Return False
            End If
            If (edt_password.Text.Trim().Length < 5) Then
                MessageBox.Show("Mật khẩu thành viên không được nhỏ hơn 5 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_password
                Return False
            End If
        End If
        'Bắt điều kiện trùng tên đăng nhập của thành viên - Username
        If (FlagEvent = 1) Then
            strSQL = String.Format("Select * From HT_ThanhVien Where ten_tv = '{0}'", Globals.Find_Replace(edt_username.Text.ToString().Trim()))
        ElseIf (FlagEvent = 2) Then
            strSQL = String.Format("Select * From HT_ThanhVien Where ten_tv = '{0}' And ten_tv <> '{1}'", Globals.Find_Replace(edt_username.Text.ToString().Trim()), userName)
        End If
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    MessageBox.Show("Tên đăng nhập này đã tồn tại. Xin hãy vui nhập lại một tên khác!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_username
                    Return False
                End If
            End If
        End Using

        If (edt_email.Text.Trim() <> "") Then ' Nếu địa chỉ e-mail của cán bộ mà nhập thì kiểm tra có hợp lệ không
            If (Not Globals.IsEmail(edt_email.Text.Trim().ToString())) Then
                MessageBox.Show("Địa chỉ e-mail của thành viên không hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_email
                Return False
            End If
        End If


        Return True
    End Function

    ''' <summary>
    ''' Thực hiện load lại Nhóm thành viên nếu có cập nhật Nhóm thành viên
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReLoad()
        If (obj_grmembers.FlagUpdate = False) Then
            Return
        End If
        cb_mgroup.Items.Clear()
        arr_Groups.Clear()
        strSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Order by ma_nhom Asc")
        arr_Groups = _Globals.Bind_ComBoBox(cb_mgroup, strSQL, "---Nhóm thành viên---")
        cb_mgroup.SelectedIndex = CType(arr_Groups.IndexOf(obj_grmembers.CodeGroup.ToString()), Int32)
        obj_grmembers.Dispose()
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHT_NguoiDung_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        arr_Groups.Clear()
        cb_mgroup.Items.Clear()
        'Nếu người dùng là cấp Chi nhánh thì không load nhóm thành viên cấp TW
        If Microsoft.VisualBasic.Left(Globals.Group, 1).ToUpper <> "U" Then
            strSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Order by ma_nhom Asc")
        Else
            strSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Where LEFT(ma_nhom,1) = 'U' Order by ma_nhom Asc")
        End If
        arr_Groups = _Globals.Bind_ComBoBox(cb_mgroup, strSQL, "---Nhóm thành viên---")
        ResetAll_Controls()
        If Not (Globals.IsIntersect(";45;46;47;48;", Globals.Roles)) Then
            btn_grmember.Enabled = False
        End If
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        ResetAll_Controls()
        If Not (Globals.IsIntersect(";45;46;47;48;", Globals.Roles)) Then
            btn_grmember.Enabled = False
        End If
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Close()
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        Try
            If (IsValid()) Then
                Select Case FlagEvent
                    Case 1
                        Dim obj_member As clsHeThong.Members = New clsHeThong.Members()
                        obj_member.id_nhom = IIf(arr_Groups.Count > 0, arr_Groups(cb_mgroup.SelectedIndex), "")
                        obj_member.ho_ten = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_fullname.Text.Trim.ToString()))
                        obj_member.ten_tv = Globals.Find_Replace(edt_username.Text.Trim.ToString())
                        obj_member.mat_khau = System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Globals.Find_Replace(edt_password.Text.Trim.ToString()), "MD5")
                        obj_member.email = Globals.Find_Replace(edt_email.Text.Trim.ToString())
                        obj_member.ghi_chu = Globals.Find_Replace(edt_note.Text.Trim.ToString())
                        obj_member.ngay_sua = DateTime.Now
                        obj_member.ngay_tao = DateTime.Now
                        _Systems.Insert_Members(obj_member)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Case 2
                        Dim obj_member As clsHeThong.Members = New clsHeThong.Members()
                        obj_member.id_nhom = IIf(arr_Groups.Count > 0, arr_Groups(cb_mgroup.SelectedIndex), "")
                        obj_member.ho_ten = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_fullname.Text.Trim.ToString()))
                        obj_member.ten_tv = Globals.Find_Replace(edt_username.Text.Trim.ToString())
                        obj_member.mat_khau = System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Globals.Find_Replace(edt_password.Text.Trim.ToString()), "MD5")
                        obj_member.email = Globals.Find_Replace(edt_email.Text.Trim.ToString())
                        obj_member.ghi_chu = Globals.Find_Replace(edt_note.Text.Trim.ToString())
                        obj_member.ngay_sua = DateTime.Now
                        obj_member.ngay_tao = DateTime.Now
                        _Systems.Update_Members(obj_member, userName)
                        If (userName = Globals.UserVal) Then
                            Globals.UserVal = Globals.Find_Replace(edt_username.Text.Trim.ToString())
                        End If
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                End Select
            End If
        Catch ex As Exception
            MessageBox.Show("Cập nhật thông tin thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub frmHT_NguoiDung_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Close()
        End If
    End Sub

    Private Sub frmHT_NguoiDung_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub btn_grmember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_grmember.Click
        obj_grmembers = New frmHT_NhomThVien
        If (arr_Groups.Count > 0) Then
            If (cb_mgroup.SelectedIndex > 0) Then
                obj_grmembers.CodeGroup = arr_Groups(cb_mgroup.SelectedIndex).ToString()
            Else
                obj_grmembers.CodeGroup = ""
            End If
        Else
            obj_grmembers.CodeGroup = ""
        End If
        obj_grmembers.Progress_Changed = New frmHT_NhomThVien.ProgressChangedEventHandler(AddressOf ReLoad)
        obj_grmembers.ShowDialog()
    End Sub
#End Region

#Region "---> Events: Các sự kiện ngoại lệ người dùng <---"
    Private Sub edt_fullname_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_fullname.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_fullname.Text.Trim() <> "") Then
                edt_username.Focus()
            Else
                edt_fullname.Focus()
            End If
        End If
    End Sub

    Private Sub edt_fullname_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_fullname.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_mgroup.Focus()
        End If
    End Sub

    Private Sub edt_username_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_username.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_username.Text.Trim() <> "") Then
                edt_password.Focus()
            Else
                edt_username.Focus()
            End If
        End If
    End Sub

    Private Sub edt_username_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_username.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_fullname.Focus()
        End If
    End Sub

    Private Sub edt_password_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_password.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_password.Text.Trim() <> "") Then
                edt_email.Focus()
            Else
                edt_password.Focus()
            End If
        End If
    End Sub

    Private Sub edt_password_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_password.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_username.Focus()
        End If
    End Sub

    Private Sub edt_email_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_email.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_note.Focus()
        End If
    End Sub

    Private Sub edt_email_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_email.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_password.Focus()
        End If
    End Sub

    Private Sub edt_note_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_note.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_note_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_note.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_email.Focus()
        End If
    End Sub
#End Region

End Class