Public Class frmHT_Login

#Region "--->Khai báo thuộc tính và khởi tạo đối tượng<---"
    ''' <summary>
    ''' Get-Set Giá trị Tên đăng nhập của thành viên
    ''' </summary>
    ''' <remarks></remarks>
    Private _UserName As String
    Public Property UserName() As String
        Get
            Return _UserName
        End Get
        Set(ByVal value As String)
            _UserName = value
        End Set
    End Property

    ''' <summary>
    ''' Get-Set Giá trị tập quyền của thành viên đăng nhập
    ''' </summary>
    ''' <remarks></remarks>
    Private _Permits As String
    Public Property Permits() As String
        Get
            Return _Permits
        End Get
        Set(ByVal value As String)
            _Permits = value
        End Set
    End Property

    ''' <summary>
    ''' Get - Set giá trị xác định admin. Nếu là True -> Là admin. False -> Là Member
    ''' </summary>
    ''' <remarks></remarks>
    Private _Admin As Boolean
    Public Property Admin() As Boolean
        Get
            Return _Admin
        End Get
        Set(ByVal value As Boolean)
            _Admin = value
        End Set
    End Property

    Private _Systems As clsHeThong = New clsHeThong
    Private _Globals As Globals = New Globals
    Dim _DBAccess As DBAccess = New DBAccess()
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
#End Region

#Region "--->Events: Các sự kiện chính<---"
    Private Sub frmHT_Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim _cfg As New clsSettings
        edt_password.Text = ""
        If _cfg.ReadValue("LuuThongTin") = "1" Then
            edt_username.Text = _cfg.ReadValue("Username")
            edt_maPOS.Text = _cfg.ReadValue("POS")
            chkLuuThongTin.Checked = True
        Else
            edt_username.Text = ""
            edt_maPOS.Text = ""
            chkLuuThongTin.Checked = True
        End If
        AddHandler edt_username.GotFocus, AddressOf edt_GotFocus
        AddHandler edt_password.GotFocus, AddressOf edt_GotFocus
        AddHandler edt_maPOS.GotFocus, AddressOf edt_GotFocus
        edt_username.Focus()
    End Sub

    Private Sub btn_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_reset.Click
        edt_password.Text = ""
        edt_username.Text = ""
        edt_maPOS.Text = ""
        ActiveControl = edt_username
    End Sub

    Private Sub btn_login_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_login.Click
        'If Not _DBAccess.CheckConnection() Then
        '    MessageBox.Show("Thông tin thiết lập kết nối cơ sở dữ liệu không hợp lệ.", "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        '    Dispose()
        '    Application.Exit()
        '    Return
        'End If

        'Kiểm tra thông tin đăng nhập có chính xác không, gán giá trị _UserName, DONVI, HoTen từ trong _Systems.IsLogin
        Dim IsValid As Byte = _Systems.IsLogin(Globals.Find_Replace(edt_username.Text.Trim), Globals.Find_Replace(edt_password.Text), Globals.Find_Replace(edt_maPOS.Text.Trim))
        Select Case IsValid
            Case 0
                ActiveControl = edt_username
                edt_username.SelectAll()
            Case 1
                ActiveControl = edt_password
                edt_password.SelectAll()
            Case 2
                _UserName = Globals.Find_Replace(edt_username.Text.Trim().ToString())
                gUsername = _UserName 'Lưu biến Username để sử dụng khi thay đổi pass
                'zzzzzzzzzzzzzzzzzzzz(
                DONVI = edt_maPOS.Text.Trim
                'zzzzzzzzzzzzzzzzzzzz)
                'Lấy chuỗi quyền của thành viên đăng nhập chương trình
                Globals.Roles = _Systems.GetRoles(_UserName)
                Globals.Group = _Systems.GetGroup(_UserName)
                Globals.QuyenQuanLyCN = _Systems.Get_QuyenQuanLyCN(_UserName)
                If (Globals.Roles = "") Then
                    My_MessageBox("Thành viên đăng nhập hiện chưa được phân quyền thao tác chương trình!")
                    Dispose()
                    Application.Exit()
                    'ElseIf Globals.QuyenQuanLyCN.Trim = "" Then
                    '    My_MessageBox("Thành viên đăng nhập hiện chưa được phân quyền thao tác với đơn vị nào!")
                    '    Dispose()
                    '    Application.Exit()
                End If

                'Lưu thông tin đăng nhập
                Dim _cfg As New clsSettings
                If chkLuuThongTin.Checked Then
                    _cfg.WriteValue("LuuThongTin", "1")
                    _cfg.WriteValue("Username", edt_username.Text.Trim)
                    _cfg.WriteValue("POS", edt_maPOS.Text.Trim)
                Else
                    _cfg.WriteValue("LuuThongTin", "0")
                    _cfg.WriteValue("Username", "")
                    _cfg.WriteValue("POS", "")
                End If
                _cfg.Save()

                DialogResult = Windows.Forms.DialogResult.OK

                'Dim dirApp As String = Application.StartupPath
                'If FileIO.FileSystem.FileExists(dirApp & "\" & "UDApp.exe") Then
                '    FileIO.FileSystem.DeleteFile(dirApp & "\" & "updateApp.exe")
                '    FileIO.FileSystem.RenameFile(dirApp & "\" & "UDApp.exe", "updateApp.exe")
                'End If
            Case Else
        End Select
    End Sub
#End Region

#Region "--->Events: Các sự kiện trợ giúp người dùng<---"
    Private Sub frmHT_Login_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.Alt = True And e.KeyCode = Keys.D) Then
                btn_login_Click(sender, Nothing)
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub edt_GotFocus(sender As Object, e As System.EventArgs) ' Handles edt_password.GotFocus
        CType(sender, TextBox).SelectAll()
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

    Private Sub edt_password_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_password.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                edt_maPOS.Focus()
            Case Keys.Up
                edt_username.Focus()
        End Select
    End Sub

    Private Sub edt_maPOS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_maPOS.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If (edt_username.Text.Trim() <> "") Then
                    btn_login_Click(sender, Nothing)
                Else
                    edt_password.Focus()
                End If
            Case Keys.Tab, Keys.Down
                btn_login.Focus()
            Case Keys.Up
                edt_username.Focus()
        End Select
    End Sub

    Private Sub frmHT_Login_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub
#End Region

End Class