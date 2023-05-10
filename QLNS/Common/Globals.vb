Imports System.Text.RegularExpressions
Imports log4net
Imports System.IO
Public Class Globals
    
    Private _SqlHelper As DBAccess

    ''' <summary>
    ''' Biến lưu tên đăng nhập của thành viên đang thao tác chương trình
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared UserVal As String = ""

    Public Shared Group As String = ""

    ''' <summary>
    ''' Biến lưu chuỗi quyền của thành viên đăng nhập chương trình
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Roles As String = ""

    ''' <summary>
    ''' Biến lưu danh sách Chi nhánh mà user này có được quyền quản lý
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared QuyenQuanLyCN As String = ""

    ''' <summary>
    ''' Khai báo biến định dạng số
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared cultureNum As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-us")

    ''' <summary>
    ''' Biến lưu danh sách các sự kiện người dùng
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum eventFlag
        viewRecord = 0
        addRecord = 1
        editRecord = 2
        deleteRecord = 3
    End Enum

    ''' <summary>
    ''' Hàm trả về giá trị thiết lập cấu hình trong file config
    ''' </summary>
    ''' <param name="strSetting">Tên của biến cấu hình trong config</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAppSetting(ByVal strSetting As String) As String
        Dim strResult = ""
        Try
            strResult = System.Configuration.ConfigurationManager.AppSettings(strSetting).ToString().Trim()
        Catch ex As Exception
            MessageBox.Show("Lấy giá trị thiết lập cấu hình: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
        Return strResult
    End Function

    Public Shared Function GetDateTime_ForServerDB() As DateTime
        Try
            Dim dDate As DateTime = DateTime.Now
            Dim sSQLDate As String = "Select dbo.GetDate_Server() CurrentDate"
            Using db_date As DataTable = SoftSqlHelper.ExecuteForSQL(sSQLDate)
                If Not (db_date Is Nothing) Then
                    If db_date.Rows.Count > 0 Then
                        dDate = Convert.ToDateTime(db_date.Rows(0)(0))
                    End If
                End If
            End Using
            Return dDate
        Catch ex As Exception
            Return DateTime.MinValue
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện chuẩn hoá chuỗi dữ liệu truyền vào
    ''' </summary>
    ''' <param name="wstr">Chuỗi cần chuẩn hoá</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Find_Replace(ByVal wstr As String) As String
        Dim _result As String = ""
        If Not IsNothing(wstr) Then
            If (wstr <> "") Then
                wstr = wstr.Trim()
                'Chuẩn hoá tiếp các trường hợp đặc biệt
                If (wstr.IndexOf("'") >= 0) Then
                    _result = wstr.Replace("'", "''")
                End If
                If (wstr.IndexOf("\\") >= 0) Then
                    _result = wstr.Replace("\\", "\\\\")
                End If
                If (wstr = "NULL") Then
                    _result = "NULL "
                Else
                    _result = wstr
                End If
                'Thực hiện xoá các khoảng cách thừa trong chuỗi
                If (_result <> "") Then
                    While (_result.IndexOf("  ") >= 0)
                        _result = _result.Replace("  ", " ")
                    End While
                End If
            End If
        End If
        Return _result
    End Function

    ''' <summary>
    ''' Hàm thực hiện xóa ký tự ở đầu, ở cuối và xóa hai ký tự sát nhau theo tham số ký tự cần xóa truyền vào 
    ''' </summary>
    ''' <param name="pInput">Chuỗi cần chuẩn hóa xóa ký tự thừa ở đầu, cuối và hai ký tự liền kề</param>
    ''' <param name="pDelimiter">Ký tự cần xóa ở đầu và cuối chuỗi và hai ký tự sát nhau</param>
    ''' <returns>Chuỗi giá trị đã được chuẩn hóa</returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChar_FirstAndLast(ByVal pInput As String, ByVal pDelimiter As String) As String
        Dim _StandardizedRet As String = ""
        If (Not String.IsNullOrEmpty(pInput)) Then
            _StandardizedRet = pInput
            _StandardizedRet = _StandardizedRet.Replace(pDelimiter + pDelimiter, pDelimiter)
            If (Not String.IsNullOrEmpty(_StandardizedRet)) Then
                While (_StandardizedRet.EndsWith(pDelimiter))       'Bỏ dấu chấm phẩy ở cuối chuỗi đi
                    _StandardizedRet = _StandardizedRet.Substring(0, _StandardizedRet.Length - 1)
                End While
            End If
            If (Not String.IsNullOrEmpty(_StandardizedRet)) Then
                While (Not String.IsNullOrEmpty(_StandardizedRet.Trim()) And _StandardizedRet.Trim().Substring(0, 1) = pDelimiter)
                    _StandardizedRet = _StandardizedRet.Trim().Substring(1)
                End While

            End If
        End If
        Return _StandardizedRet
    End Function


    ''' <summary>
    ''' Hàm thực hiện chuyển hết cac ký tự đầu của một chuỗi thành Ký tự Hoa
    ''' </summary>
    ''' <param name="_strval">Chuỗi truyền vào</param>
    ''' <returns>Chuỗi đã đc chuẩn hoá</returns>
    ''' <remarks></remarks>
    Public Shared Function MakeFirstWordCharUpper(ByVal _strval As String) As String
        Dim _result As String = _strval.Trim().ToLower()
        If (_result.Length <= 1) Then Return _result.ToUpper()
        Dim arrLetters As Char() = _result.ToCharArray()
        arrLetters(0) = Char.ToUpper(arrLetters(0))
        For i As Integer = 1 To _result.Length - 1
            If (arrLetters(i - 1) = Chr(32)) Then
                arrLetters(i) = Char.ToUpper(arrLetters(i))
            End If
        Next
        Return New String(arrLetters)
    End Function

    ''' <summary>
    ''' Hàm kiểm tra xem địa chỉ e-mail có hợp lệ không
    ''' </summary>
    ''' <param name="strEmail">Địa chỉ e-mail cần kiểm tra</param>
    ''' <returns>True: Hợp lệ (Valid). False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmail(ByVal strEmail As String) As Boolean
        Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" & _
                                 "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" & _
                                 ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
        Dim _rex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
        If (_rex.IsMatch(strEmail)) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Method to determine is the absolute file path is a valid path
    ''' </summary>
    ''' <param name="strPath">The path we want to check</param>
    ''' <returns>True: Hợp lệ (Valid). False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Public Shared Function IsValidPath(ByVal strPath As String) As Boolean
        Dim strRegex As String = "^(([a-zA-Z]\:)|(\\))(\\{1}|((\\{1})[^\\]([^/:*?<>""|]*))+)$"
        Dim _rex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex, System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If (_rex.IsMatch(strPath)) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Hàm thực hiện định dạng lại chuỗi trong textbox đúng kiểu số thực
    ''' </summary>
    ''' <param name="edt_name">Tên textbox cần định dạng</param>
    ''' <param name="_NFInfo">Kiểu định dạng số</param>
    ''' <returns>Số thực sau khi đã định dạng</returns>
    ''' <remarks></remarks>
    Public Shared Function Convert_ToDouble(ByVal edt_name As TextBox, ByVal _NFInfo As System.Globalization.NumberFormatInfo) As Double
        Try
            Return Double.Parse(edt_name.Text.Trim(), _NFInfo)
        Catch
            edt_name.Text = "0"
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện kiểm tra một chuỗi truyền vào có phải là số không ?
    ''' </summary>
    ''' <param name="strValue">Giá trị cần kiểm tra</param>
    ''' <returns>True: Là số. False: Không phải là số</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNumeric(ByVal strValue As String) As Boolean
        If (strValue <> "") Then
            Try
                Double.Parse(strValue)
            Catch
                Return False
            End Try
            Return True
        Else : Return False
        End If
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra một chuỗi truyền vào có phải là Datetime không ?
    ''' </summary>
    ''' <param name="strValue">Giá trị cần kiểm tra</param>
    ''' <returns>True: Là Datetime. False: Không phải là Datetime</returns>
    ''' <remarks></remarks>
    Public Shared Function IsDatetime(ByVal strValue As String) As Boolean
        If (strValue.Trim().Length > 0) Then
            Try
                DateTime.Parse(strValue)
                Return True
            Catch
                Return False
            End Try
        Else : Return False
        End If
    End Function

    ''' <summary>
    ''' Hàm trả về tên của sự kiện người dùng - Dựa vào chỉ số sự kiện
    ''' </summary>
    ''' <param name="_EventId">Chỉ số xác định sự kiện</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetName_Events(ByVal _EventId As Byte) As String
        Select Case _EventId
            Case 1
                Return "thêm mới"
            Case 2
                Return "sửa đổi"
            Case 3
                Return "xoá bỏ"
            Case Else
                Return ""
        End Select
    End Function

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào ComboBox. Functions is fill data in combobox 
    ''' </summary>
    ''' <param name="cb_name">Tên combo cần đổ dl</param>
    ''' <param name="strSQL">Câu lệnh truy xuất dl</param>
    ''' <param name="strNone">Dòng text đầu tiên lựa chọn combobox</param>
    ''' <returns>Mảng lưu chỉ số các phần tử dl</returns>
    ''' <remarks></remarks>
    Public Function Bind_ComBoBox(ByVal cb_name As ComboBox, ByVal strSQL As String, ByVal strNone As String) As ArrayList
        '_SqlHelper = New DBAccess
        Dim arrList As ArrayList = New ArrayList
        arrList.Clear()
        Dim db As DataTable = New DataTable
        cb_name.Items.Clear()
        If (strSQL <> "") Then
            'db = _SqlHelper.getDataTable(strSQL)
            db = SoftSqlHelper.ExecuteForSQL(strSQL)
            Dim dr As DataRow = db.NewRow()
            dr(0) = 0
            dr(1) = strNone
            db.Rows.InsertAt(dr, 0)
            Dim i As Integer = 0
            Dim rows As Int32 = db.Rows.Count - 1
            For i = 0 To rows
                cb_name.Items.Add(db.Rows(i)(1).ToString())
                arrList.Add(IIf(db.Rows(i)(0).ToString() <> "", db.Rows(i)(0).ToString(), ""))
            Next
            If (cb_name.Items.Count <> 0) Then
                Try
                    cb_name.SelectedIndex = 0
                Catch ex As Exception

                End Try
            End If
        End If
        Return arrList
    End Function

    Public Function BindList_ComBoBox_ListBox(ByVal cb_name As ComboBox, ByVal lb_name As ListBox, ByVal pQuerySQL As String, ByVal pNone As String, pFlagCall As Byte) As ArrayList
        Dim ARL_ListCodes As ArrayList = New ArrayList
        ARL_ListCodes.Clear()
        Dim db As DataTable = New DataTable
        If pFlagCall = 1 Then
            cb_name.Items.Clear()
            If (pQuerySQL <> "") Then
                db = SoftSqlHelper.ExecuteForSQL(pQuerySQL)
                Dim dr As DataRow = db.NewRow()
                dr(0) = ""
                dr(1) = pNone
                db.Rows.InsertAt(dr, 0)
                Dim i As Integer = 0
                Dim rows As Int32 = db.Rows.Count - 1
                For i = 0 To rows
                    cb_name.Items.Add(db.Rows(i)(1).ToString())
                    ARL_ListCodes.Add(IIf(db.Rows(i)(0).ToString() <> "", db.Rows(i)(0).ToString(), ""))
                Next
                If (cb_name.Items.Count <> 0) Then
                    Try
                        cb_name.SelectedIndex = 0
                    Catch ex As Exception

                    End Try
                End If
            End If
        End If
        Return ARL_ListCodes
    End Function

    ''' <summary>
    ''' Hàm thực hiện Fill dữ liệu vào ListBox
    ''' </summary>
    ''' <param name="lb_name">Tên listbox cần fill dữ liệu vào</param>
    ''' <param name="strSQL">Câu lệnh SQL để lấy dữ liệu</param>
    ''' <param name="strNone">Chuỗi chú thích - chính là item 0</param>
    ''' <returns>Mảng danh sách các bản ghi</returns>
    ''' <remarks></remarks>
    Public Function Bind_ListBox(ByVal lb_name As ListBox, ByVal strSQL As String, ByVal strNone As String) As ArrayList
        _SqlHelper = New DBAccess
        Dim arrList As ArrayList = New ArrayList
        arrList.Clear()
        Dim db As DataTable = New DataTable
        lb_name.Items.Clear()
        If (strSQL <> "") Then
            db = _SqlHelper.SelectDBRows(strSQL)
            Dim dr As DataRow = db.NewRow()
            dr(0) = 0
            dr(1) = strNone
            db.Rows.InsertAt(dr, 0)
            Dim rows As Int32 = db.Rows.Count - 1
            For i As Integer = 0 To rows
                lb_name.Items.Add(db.Rows(i)(1).ToString())
                arrList.Add(IIf(db.Rows(i)(0).ToString() <> "", db.Rows(i)(0).ToString(), ""))
            Next
            If (lb_name.Items.Count <> 0) Then
                lb_name.SelectedIndex = 0
            End If
        End If
        Return arrList
    End Function

    Public Function Bind_ListBox(ByVal lb_name As ListBox, ByVal strSQL As String) As ArrayList
        _SqlHelper = New DBAccess
        Dim arrList As ArrayList = New ArrayList
        arrList.Clear()
        lb_name.Items.Clear()
        If (strSQL <> "") Then
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        Dim rows As Int32 = db.Rows.Count - 1
                        For i As Int32 = 0 To rows
                            lb_name.Items.Add(db.Rows(i)(1).ToString())
                            arrList.Add(IIf(db.Rows(i)(0).ToString() <> "", db.Rows(i)(0).ToString(), ""))
                        Next
                    End If
                End If
            End Using
            If (lb_name.Items.Count <> 0) Then
                lb_name.SelectedIndex = 0
            End If
        End If
        Return arrList
    End Function

    ''' <summary>
    ''' Hàm thực hiện tìm kiếm một TreeNode trong một treeview
    ''' </summary>
    ''' <param name="_nodes"></param>
    ''' <param name="_code"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function TreeViewFindNode(ByVal _nodes As TreeNodeCollection, ByVal _code As String) As TreeNode
        Dim _result As String = ""
        If (_nodes Is Nothing Or _nodes.Count = 0 Or _code = String.Empty) Then
            Return Nothing
        End If

        For Each node As TreeNode In _nodes
            If Not (node Is Nothing) Then
                _result = node.Tag.ToString()
                If (_result = _code) Then
                    Return node
                End If
            End If
            'Current node isn't the one so search it's children (there may be none)
            Dim foundNode As TreeNode = TreeViewFindNode(node.Nodes, _code)
            'If it was found in the child collection return it
            If Not (foundNode Is Nothing) Then
                Return foundNode
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm thực hiện Checked tất các các items trên lưới dữ liệu khi Checked một CheckBox
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="status">True: Checked False: UnChecked</param>
    ''' <remarks></remarks>
    Public Shared Sub Check_All_Items(ByVal dgv_name As DataGridView, ByVal status As Boolean)
        If (dgv_name.Rows.Count < 1) Then
            Return
        Else
            For i As Integer = 0 To dgv_name.Rows.Count - 1
                dgv_name.Rows(i).Cells("cln_Choice").Value = status
            Next
        End If
    End Sub

    ''' <summary>
    ''' Class liên quan đến file. Write by: Dương Văn Chữ - Date: 01-08-2008
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FileDao

        ''' <summary>
        ''' Hàm thực hiện kiểm tra đường dẫn thư mục có tồn tại không - Returns True if the provided path is an existing directory
        ''' </summary>
        ''' <param name="strPath">Đường dẫn thư mục cần kiểm tra</param>
        ''' <returns>True: Đường dẫn tồn tại - False: Đường dẫn không tồn tại</returns>
        ''' <remarks></remarks>
        Public Shared Function IsDirectory(ByVal strPath As String) As Boolean
            Try
                Dim _result As Boolean = False
                _result = System.IO.Directory.Exists(strPath)
                Return _result
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Hàm thực hiện kiểm tra đường dẫn file có tồn tại không - Returns True if the provided path is an existing file
        ''' </summary>
        ''' <param name="strPath">Đường dẫn file cần kiểm tra</param>
        ''' <returns>True: Nếu đường dẫn tồn tại. False: Đường dẫn không tồn tại</returns>
        ''' <remarks></remarks>
        Public Shared Function IsFile(ByVal strPath As String) As Boolean
            Return System.IO.File.Exists(strPath)
        End Function

        ''' <summary>
        ''' Định nghĩa cấu trúc của file
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure RowFile
            Public Id As Int32
            Public Path As String
            Public FileName As String
            Public FileSize As String
            Public FormatFile As String
            Public Sub New(ByVal _Id As Int32, ByVal _Path As String, ByVal _FileName As String, ByVal _Sizefile As String, ByVal _Formatfile As String)
                Id = _Id
                Path = _Path
                FileName = _FileName
                FileSize = _Sizefile
                FormatFile = _Formatfile
            End Sub
        End Structure

        ''' <summary>
        ''' Hàm trả về tên file - Từ đường dẫn truyền vào
        ''' </summary>
        ''' <param name="strfile">Đường dẫn của file --> Cần lấy tên file</param>
        ''' <param name="IsOK">True: Tên file không có đuôi. False: Tên file gồm cả đuôi file</param>
        ''' <returns>Tên của file</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFileName(ByVal strfile As String, ByVal IsOK As Boolean) As String
            Dim n As Int32 = strfile.LastIndexOf("\")
            strfile = strfile.Substring(n + 1)
            If (IsOK = True) Then
                strfile = strfile.Substring(0, strfile.LastIndexOf("."))
            End If
            Return strfile
        End Function

        ''' <summary>
        ''' Hàm trả về duôi của file - Từ đường dẫn file truyền vào
        ''' </summary>
        ''' <param name="strPathfile">Đường dẫn của file</param>
        ''' <returns>Trả về đuôi của file</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFormatSize(ByVal strPathfile As String) As String
            Dim _result As String = ""
            If (strPathfile <> "") Then
                If (System.IO.File.Exists(strPathfile)) Then
                    Dim _fileInfo As System.IO.FileInfo = New System.IO.FileInfo(strPathfile)
                    _result = _fileInfo.Extension.ToString()
                End If
            End If
            Return _result
        End Function

        ''' <summary>
        ''' Hàm thực hiện copy file
        ''' </summary>
        ''' <param name="sourceFile">Đường dẫn chứa file nguồn</param>
        ''' <param name="destFile">Đường dẫn file đích</param>
        ''' <returns>True: copy thành công. False: Không thành công</returns>
        ''' <remarks></remarks>
        Public Shared Function CopyFile(ByVal sourceFile As String, ByVal destFile As String) As Boolean
            Try
                System.IO.File.Copy(sourceFile, destFile, True)
                Return True
            Catch ex As Exception
                MessageBox.Show("Thực hiện copy file không thành công: \n" + GetFileName(sourceFile, False), "Copy file ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Hàm thực hiện xoá file đính kèm
        ''' </summary>
        ''' <param name="strPath">Đường dẫn file cần xoá</param>
        ''' <remarks></remarks>
        Public Shared Sub DeleteFile(ByVal strPath As String)
            Try
                If (System.IO.File.Exists(strPath)) Then
                    System.IO.File.Delete(strPath)
                End If
            Catch ex As Exception
                MessageBox.Show("Tháo tác không hợp lệ. Có thể file này đang được sử dụng!", "Delete file ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            End Try
        End Sub

        ''' <summary>
        ''' Hàm tạo folder trong đường dẫn truyền vào - Create a subfolder in the current path
        ''' </summary>
        ''' <param name="strPath">Path current folder</param>
        ''' <param name="folderName">Folder name</param>
        ''' <returns>True: Tạo thành công thư mục. False: Không thành công</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeFolder(ByVal strPath As String, ByVal folderName As String) As Boolean
            strPath = System.IO.Path.Combine(strPath, folderName)
            Try
                If (Not (IsDirectory(strPath))) Then
                    System.IO.Directory.CreateDirectory(strPath)
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                MessageBox.Show("Error: " + ex.Message.ToString(), "Create folder ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Hàm thực hiện tạo thư mục theo đường dẫn truyền vào
        ''' </summary>
        ''' <returns>True: Tạo thành công. False: Tạo không thành công</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeFolder(ByVal _pathFolder As String) As Boolean
            If (_pathFolder.Substring(_pathFolder.Length - 1) <> "\") Then
                _pathFolder = _pathFolder + "\"
            End If
            Try
                If (Not (Globals.FileDao.IsDirectory(_pathFolder))) Then
                    System.IO.Directory.CreateDirectory(_pathFolder)
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                MessageBox.Show("Error: " + ex.Message.ToString(), "Create folder ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Hàm trả về số lượng file có trong một thư mục truyền vào
        ''' </summary>
        ''' <param name="dirPath">Thư mục cần xem có bao nhiêu file trong đó</param>
        ''' <returns>Số lượng file trong thư mục</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCountFiles(ByVal dirPath As String) As Int16
            Dim _count As Int16 = 0
            If (IO.Directory.Exists(dirPath)) Then
                Dim dirInfor As New IO.DirectoryInfo(dirPath)
                Dim fileInfor As IO.FileInfo() = dirInfor.GetFiles("*.*")
                _count = fileInfor.Length - 1
            End If
            Return _count
        End Function

        ''' <summary>
        ''' Hàm thực hiện - Xoá thư mục và file hoặc subfolder trong thư mục theo đường dẫn truyền vào
        ''' </summary>
        ''' <param name="strPath">Đường dẫn thư mục cần delete</param>
        ''' <returns>True: Thành công. False: Không thành công</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteFolder(ByVal strPath) As Boolean
            Try
                If Not (System.IO.Directory.Exists(strPath)) Then
                    Return False
                Else
                    Dim dirInfo As New System.IO.DirectoryInfo(strPath)
                    'Xoá các file trong đường dẫn thư mục truyền vào
                    For Each _file As IO.FileInfo In dirInfo.GetFiles("*.*")
                        System.IO.File.Delete(System.IO.Path.Combine(strPath, _file.Name))
                    Next
                    'Thực hiện xoá các subfolder trong đường dẫn thư mục truyền vào
                    For Each _subdir As IO.DirectoryInfo In dirInfo.GetDirectories()
                        DeleteFolder(System.IO.Path.Combine(strPath, _subdir.Name))
                    Next
                    System.IO.Directory.Delete(strPath, True)
                    Return True
                End If
            Catch ex As Exception
                MessageBox.Show("Xoá thư mục: " + ex.Message.ToString(), "Delete folder ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return False
            End Try
        End Function

    End Class

    'Các hàm liên quan đến CheckedListBox
    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào CheckedListBox
    ''' </summary>
    ''' <param name="clb_name">Tên CheckedListBox</param>
    ''' <param name="strSQL">Câu lệnh truy xuất dữ liệu</param>
    ''' <returns>Mảng chứa danh sách id dữ liệu fill vào CheckedListBox</returns>
    ''' <remarks></remarks>
    Public Function FillData_CheckedListBox(ByVal clb_name As CheckedListBox, ByVal strSQL As String) As ArrayList
        _SqlHelper = New DBAccess
        Dim arrList As ArrayList = New ArrayList
        arrList.Clear()
        clb_name.Items.Clear()
        If (strSQL <> "") Then
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        For i As Integer = 0 To db.Rows.Count - 1
                            clb_name.Items.Add(db.Rows(i)(1).ToString())
                            arrList.Add(db.Rows(i)(0).ToString())
                        Next
                    End If
                End If
            End Using
        End If
        Return arrList
    End Function

    ''' <summary>
    ''' 'Hàm trả về Chuỗi các id đã checked trong CheckedListBox
    ''' </summary>
    ''' <param name="clb_Name">Name CheckedListBox</param>
    ''' <returns>Chuỗi id</returns>
    ''' <remarks></remarks>
    Public Function GetItems_CheckedListBox(ByVal clb_Name As CheckedListBox, ByVal arrList As ArrayList) As String
        Dim result As String = ""
        For Each indexChecked As Int32 In clb_Name.CheckedIndices
            result = result + arrList(indexChecked).ToString() + ";"
        Next
        Return result
    End Function

    ''' <summary>
    ''' Hàm thực thiện Checked các items tương ứng theo chuỗi id
    ''' </summary>
    ''' <param name="clb_name">CheckedListBox</param>
    ''' <param name="strId">Chuỗi id</param>
    ''' <param name="arrId">Mảng lưu danh sách id</param>
    ''' <remarks></remarks>
    Public Sub SetItems_CheckedListBox(ByVal clb_name As CheckedListBox, ByVal strId As String, ByVal arrId As ArrayList)
        'Thực hiện bỏ dấu (;) ở đầu chuỗi nếu có
        If strId.Substring(0, 1) = ";" Then
            strId = strId.Substring(1)
        End If
        'Bắt đầu tách chỉ số quyền trong danh sách tập quyền
        Dim strTemp As String = strId
        Dim i As Integer = 0
        While (strTemp.Trim() <> "")
            i = strTemp.IndexOf(";")
            If (i > 0) Then
                Dim k As Integer = CType(arrId.IndexOf(strTemp.Substring(0, i)), Integer)
                clb_name.SetItemChecked(k, True)
                strTemp = strTemp.Substring(i + 1)
            Else
                strTemp = ""
            End If
        End While
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Reset các checked của các items trong CheckedListBox
    ''' </summary>
    ''' <param name="clb_name">Name CheckedListBox</param>
    ''' <remarks></remarks>
    Public Sub ResetItems_CheckedListBox(ByVal clb_name As CheckedListBox)
        If (clb_name.Items.Count > 0) Then
            For i As Int32 = 0 To clb_name.Items.Count - 1
                clb_name.SetItemCheckState(i, CheckState.Unchecked)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Checked toàn bộ các items của CheckedListBox khi CheckBox được checked và ngược lại
    ''' </summary>
    ''' <param name="ckb_name">Tên CheckBox</param>
    ''' <param name="clb_name">Tên của CheckedListBox</param>
    ''' <remarks></remarks>
    Public Shared Sub CheckAll_CheckListBox(ByVal ckb_name As CheckBox, ByVal clb_name As CheckedListBox)
        If (clb_name.Items.Count > 0) Then
            If (ckb_name.Checked = True) Then
                For i As Integer = 0 To clb_name.Items.Count - 1
                    'clb_name.SetItemCheckState(i, CheckState.Checked)
                    clb_name.SetItemChecked(i, True)
                Next
            Else
                For i As Integer = 0 To clb_name.Items.Count - 1
                    'clb_name.SetItemCheckState(i, CheckState.Unchecked)
                    clb_name.SetItemChecked(i, False)
                Next
            End If
        Else
            ckb_name.Checked = False
        End If
    End Sub

    Public Shared Function Splip_Strings(ByVal _lists As String, ByVal _cdelimiter As String) As String()
        Dim ARL_Rets As String() = Nothing
        Try
            ARL_Rets = _lists.Split(New String() {_cdelimiter}, StringSplitOptions.RemoveEmptyEntries)
        Catch ex As Exception
            ARL_Rets = New String() {String.Empty}
        End Try
        Return ARL_Rets
    End Function

    ''' <summary>
    ''' Hàm thực hiện tách các phần tử của chuỗi cách nhau bởi dấu (;) và bind vào mảng
    ''' </summary>
    ''' <param name="strs">Chuỗi cần tách phần tử ra</param>
    ''' <returns>Mảng lưu các phần tử được tách ra</returns>
    ''' <remarks></remarks>
    Public Shared Function Splip_Strings(ByVal strs As String) As ArrayList
        Dim arrElement As ArrayList = New ArrayList
        If (strs <> "") Then
            If strs.Substring(0, 1) = ";" Then 'Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
                strs = strs.Substring(1)
            End If
            Dim Temp As String = strs
            While Temp.EndsWith(";") 'Bỏ dấu chấm phẩy ở cuối chuỗi đi
                Temp = Temp.Substring(0, Temp.Length - 1)
            End While
            Dim arrTemp() As String = Temp.Split(";") 'Gán chuỗi Id vào mảng
            Dim v As String = ""
            Dim lengths As Integer = arrTemp.Length - 1
            For i As Integer = 0 To lengths
                v = arrTemp(i).ToString()
                arrElement.Add(v)
            Next
        End If
        Return arrElement
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Xoá các phần tử trùng nhau trong mảng (Hàm đệ quy)
    ''' </summary>
    ''' <param name="ARL_Name">Mảng chứa các phần tử trùng nhau cần xoá</param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveItems_Exists(ByVal ARL_Name As ArrayList)
        Dim i As Integer
        For i = 0 To ARL_Name.Count - 1
            If i < ARL_Name.Count - 1 Then
                If (ARL_Name(i).ToString() = ARL_Name(i + 1).ToString()) Then
                    ARL_Name.RemoveAt(i)
                    RemoveItems_Exists(ARL_Name)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Hàm trả về số ngày ---> từ hai khoảng thời gian
    ''' </summary>
    ''' <param name="dt_begin">Thời gian đầu</param>
    ''' <param name="dt_end">Thời gian cuối</param>
    ''' <returns>Số ngày cần lấy</returns>
    ''' <remarks></remarks>
    Public Shared Function GetNumberDay(ByVal dt_begin As DateTimePicker, ByVal dt_end As DateTimePicker) As Integer
        If (dt_begin.Value <= dt_end.Value) Then
            Dim _Ts As TimeSpan = dt_end.Value - dt_begin.Value
            Return _Ts.Days
        End If
        Return 0
    End Function

    ''' <summary>
    ''' Hàm kiểm tra hai chuỗi quyền truyền vào có giao nhau không. Thực chất hàm này kiểm tra:
    ''' Từng phần tử trong chuỗi quyền _roleUser có phần tử nào thuộc (hay giống một phần tử nào đó trong) chuỗi quyền _roleModule không
    ''' </summary>
    ''' <param name="_roleModule">Chuỗi quyền của một module nào đó. Lưu ý chuỗi này phải có ký tự phân cách ở đầu và cuối (VD: ;111;112;)</param>
    ''' <param name="_roleUser">Chuỗi quyền của thành viên đăng nhập. Lưu ý chuỗi này phải có ký tự phân cách ở đầu và cuối (VD: ;111;112;)</param>
    ''' <returns>True: Có thuộc. False: Hai chuỗi không giao nhau</returns>
    ''' <remarks></remarks>
    Public Shared Function IsIntersect(ByVal _roleModule As String, ByVal _roleUser As String) As Boolean
        If (_roleModule = "" Or _roleUser = "") Then Return False
        If (_roleModule.IndexOf(";") < 0 Or _roleUser.IndexOf(";") < 0) Then Return False

        If (_roleModule.EndsWith(";") = False) Then
            _roleModule = _roleModule + ";"
        End If
        _roleUser = _roleUser.Substring(1, _roleUser.LastIndexOf(";"))
        Dim arrTemp() As String = _roleUser.Split(";")
        For Each v As String In arrTemp
            If (_roleModule.IndexOf(";" + v + ";") >= 0) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Hàm thực hiện đếm số ngày thực tế làm việc giữa hai khoảng thời gian
    ''' </summary>
    ''' <param name="dt_begin">Thời gian đầu</param>
    ''' <param name="dt_end">Thời gian sau</param>
    ''' <param name="_ReposeDays">Số ngày nghỉ (T7 và Chủ nhật) giữa 2 khoảng thời gian</param>
    ''' <returns>Số ngày (đã trừ T7 và Chủ nhật) từ dt_begin đến dt_end</returns>
    ''' <remarks></remarks>
    Public Shared Function Countworkdays(ByVal dt_begin As DateTime, ByVal dt_end As DateTime, ByRef _ReposeDays As Integer) As Integer
        Dim _BusinessdayCounter As Integer = 0
        If (dt_begin <= dt_end) Then
            Dim _Ts As TimeSpan = dt_end.Subtract(dt_begin)
            Dim dt_temp As DateTime
            For i As Integer = 0 To _Ts.Days
                dt_temp = dt_begin.AddDays(i)
                If dt_temp.DayOfWeek = DayOfWeek.Saturday Or dt_temp.DayOfWeek = DayOfWeek.Sunday Then
                    _ReposeDays += 1
                Else
                    _BusinessdayCounter += 1
                End If
            Next
        End If
        Return _BusinessdayCounter
    End Function

    ''' <summary>
    ''' Hàm thực hiện đếm số ngày thực tế làm việc giữa hai khoảng thời gian
    ''' </summary>
    ''' <param name="dt_begin">Thời gian đầu</param>
    ''' <param name="dt_end">Thời gian sau</param>
    ''' <returns>Số ngày (đã trừ T7 và Chủ nhật) từ dt_begin đến dt_end</returns>
    ''' <remarks></remarks>
    Public Shared Function Countworkdays(ByVal dt_begin As DateTime, ByVal dt_end As DateTime) As Integer
        Dim _BusinessdayCounter As Integer = 0
        If (dt_begin <= dt_end) Then
            Dim _Ts As TimeSpan = dt_end.Subtract(dt_begin)
            Dim dt_temp As DateTime
            For i As Integer = 0 To _Ts.Days
                dt_temp = dt_begin.AddDays(i)
                If (dt_temp.DayOfWeek <> DayOfWeek.Saturday And dt_temp.DayOfWeek <> DayOfWeek.Sunday) Then
                    _BusinessdayCounter += 1
                End If
            Next
        End If
        Return _BusinessdayCounter
    End Function

    ''' <summary>
    ''' Hàm thực hiện đếm số ngày thực tế làm việc giữa hai khoảng thời gian (Sử dụng hàm [dbo].[countWorkDays] viết tại CSDL)
    ''' </summary>
    ''' <param name="dt_begin">Thời gian đầu</param>
    ''' <param name="dt_end">Thời gian sau</param>
    ''' <returns>Số ngày (đã trừ T7, Chủ nhật, Ngày nghỉ lễ) từ dt_begin đến dt_end</returns>
    ''' <remarks></remarks>
    Public Shared Function GetCountWorkDays(ByVal dt_begin As DateTime, ByVal dt_end As DateTime) As Integer
        Dim _BusinessdayCounter As Integer = 0
        If (dt_begin <= dt_end) Then
            _BusinessdayCounter = SoftSqlHelper.GetNumber(String.Format("Select [dbo].[countWorkDays] ('{0}', '{1}') SoNgay", DateTimeUtil.DateTimeToString(dt_begin, "yyyy-MM-dd"), DateTimeUtil.DateTimeToString(dt_end, "yyyy-MM-dd")), 0)
        End If
        Return _BusinessdayCounter
    End Function

    'Public Shared Function IsNumeric(ByVal col As System.Data.DataColumn) As Boolean
    '    If (col Is Nothing) Then
    '        Return False
    '    End If

    '    Dim numericTypes = {GetType(Byte), GetType(Decimal), GetType(Double), GetType(Int16), GetType(Int32), GetType(Int64), GetType(SByte), GetType(Single), GetType(UInt16), GetType(UInt32), GetType(UInt64)}
    '    Return numericTypes.Contains(col.DataType)
    'End Function

    Public Shared Function IsNumeric(ByVal col As DataColumn) As Boolean
        If col Is Nothing Then Return False
        Select Case col.DataType
            Case GetType(Byte), GetType(Decimal), GetType(Double), GetType(Short), GetType(Integer), GetType(Long), GetType(SByte), GetType(Single), GetType(UShort), GetType(UInteger), GetType(ULong)
                Return True
            Case Else
                Return False
        End Select
    End Function


    ' Function to test for Positive Integers
    Public Shared Function IsNaturalNumber(ByVal strNumber As [String]) As Boolean
        Dim objNotNaturalPattern As New System.Text.RegularExpressions.Regex("[^0-9]")
        Dim objNaturalPattern As New System.Text.RegularExpressions.Regex("0*[1-9][0-9]*")
        Return Not objNotNaturalPattern.IsMatch(strNumber) AndAlso objNaturalPattern.IsMatch(strNumber)
    End Function

    ' Function to test for Positive Integers with zero inclusive
    Public Shared Function IsWholeNumber(ByVal strNumber As [String]) As Boolean
        Dim objNotWholePattern As New System.Text.RegularExpressions.Regex("[^0-9]")
        Return Not objNotWholePattern.IsMatch(strNumber)
    End Function

    ' Function to Test for Integers both Positive & Negative
    Public Shared Function IsInteger(ByVal strNumber As [String]) As Boolean
        Dim objNotIntPattern As New System.Text.RegularExpressions.Regex("[^0-9-]")
        Dim objIntPattern As New System.Text.RegularExpressions.Regex("^-[0-9]+$|^[0-9]+$")
        Return Not objNotIntPattern.IsMatch(strNumber) AndAlso objIntPattern.IsMatch(strNumber)
    End Function

    ' Function to Test for Positive Number both Integer & Real
    Public Shared Function IsPositiveNumber(ByVal strNumber As [String]) As Boolean
        Dim objNotPositivePattern As New System.Text.RegularExpressions.Regex("[^0-9.]")
        Dim objPositivePattern As New System.Text.RegularExpressions.Regex("^[.][0-9]+$|[0-9]*[.]*[0-9]+$")
        Dim objTwoDotPattern As New System.Text.RegularExpressions.Regex("[0-9]*[.][0-9]*[.][0-9]*")
        Return Not objNotPositivePattern.IsMatch(strNumber) AndAlso objPositivePattern.IsMatch(strNumber) AndAlso Not objTwoDotPattern.IsMatch(strNumber)
    End Function

    ' Function to test whether the string is valid number or not
    Public Shared Function IsNumber(ByVal strNumber As [String]) As Boolean
        Dim objNotNumberPattern As New System.Text.RegularExpressions.Regex("[^0-9.-]")
        Dim objTwoDotPattern As New System.Text.RegularExpressions.Regex("[0-9]*[.][0-9]*[.][0-9]*")
        Dim objTwoMinusPattern As New System.Text.RegularExpressions.Regex("[0-9]*[-][0-9]*[-][0-9]*")
        Dim strValidRealPattern As [String] = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$"
        Dim strValidIntegerPattern As [String] = "^([-]|[0-9])[0-9]*$"
        Dim objNumberPattern As New System.Text.RegularExpressions.Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")")
        Return Not objNotNumberPattern.IsMatch(strNumber) AndAlso Not objTwoDotPattern.IsMatch(strNumber) AndAlso Not objTwoMinusPattern.IsMatch(strNumber) AndAlso
        objNumberPattern.IsMatch(strNumber)
    End Function

    ' Function To test for Alphabets.
    Public Shared Function IsAlpha(ByVal strToCheck As [String]) As Boolean
        Dim objAlphaPattern As New System.Text.RegularExpressions.Regex("[^a-zA-Z]")
        Return Not objAlphaPattern.IsMatch(strToCheck)
    End Function

    ' Function to Check for AlphaNumeric.
    Public Shared Function IsAlphaNumeric(ByVal strToCheck As [String]) As Boolean
        Dim objAlphaNumericPattern As New System.Text.RegularExpressions.Regex("[^a-zA-Z0-9]")
        Return Not objAlphaNumericPattern.IsMatch(strToCheck)
    End Function

    Public Shared Function RemoveCharacterNonNumber(ByVal Readlinevalue As String) As String
        Readlinevalue = Readlinevalue.Replace(vbLf, String.Empty)
        Readlinevalue = Regex.Replace(Readlinevalue, "[^0-9]", String.Empty)
        'Readlinevalue = Regex.Replace(Readlinevalue, "[^A-Za-z0-9\-/]", String.Empty)
        'Dim objNotWholePattern As New System.Text.RegularExpressions.Regex("[^0-9]")
        Return Readlinevalue
    End Function

    Public Shared Logger As log4net.ILog
    Public Sub New()
        'Logger = log4net.LogManager.GetLogger(New StackFrame(1, True).GetMethod().DeclaringType.Name)
        Logger = log4net.LogManager.GetLogger("Log4NetQLNS")
    End Sub

    Private Shared locker As Object = New Object()
    Public Shared Sub WriteToLog(ByVal pLogPathFile As String, ByVal pMessage As String)
        Try
            If (Not System.IO.File.Exists(pLogPathFile)) Then
                System.IO.File.Create(pLogPathFile)
            End If

            Dim sw As StreamWriter
            sw = File.AppendText(pLogPathFile)
            sw.WriteLine(DateTime.Now + vbTab + pMessage)
            sw.Close()

            'Using tw As System.IO.TextWriter = System.IO.TextWriter.Synchronized(System.IO.File.AppendText(pLogPathFile))
            '    tw.WriteLine(DateTime.Now + vbTab + pMessage)
            '    tw.Close()
            'End Using

            'SyncLock locker
            '    Using tw As System.IO.TextWriter = System.IO.TextWriter.Synchronized(System.IO.File.AppendText(pLogPath))
            '        tw.WriteLine(DateTime.Now + vbTab + pMessage)
            '    End Using
            'End SyncLock
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
    End Sub
    'Vẫn đang lỗi
    Public Shared Sub LogError(ByVal obj As Object, ByVal method As System.Reflection.MethodBase, ByVal ex As Exception)
        Dim depth As Integer = 5
        Dim message As String = IIf(String.IsNullOrEmpty(ex.Message), String.Empty, ex.Message)

        Dim innerEx As Exception = ex.InnerException
        While (innerEx IsNot Nothing AndAlso innerEx.Message IsNot Nothing AndAlso depth > 0)
            depth -= 1
            message = String.Format("{0}-{1}", message, innerEx.Message)
            innerEx = innerEx.InnerException
        End While

        Dim paramString As String = String.Empty
        Dim param = method.GetParameters()
        For i As Integer = 0 To param.Count() - 1
            If i = 0 Then
                paramString = String.Format("{0} {1}", param(i).ParameterType.FullName, param(i).Name)
            Else
                paramString = String.Format("{0}, {1} {2}", paramString, param(i).ParameterType.FullName, param(i).Name)
            End If
        Next
        Logger.Error(String.Format("{0}({1}): {2}", method.Name, paramString, message))
        'Logger.Error(obj, String.Format("{0}({1}): {2}", method.Name, paramString, message))
        'Log.Error(obj, string.Format("{0}({1}): {2}", method.Name, paramString, message));
        'Globals.Logger.Error("Lỗi chương trình")
    End Sub
    
End Class
