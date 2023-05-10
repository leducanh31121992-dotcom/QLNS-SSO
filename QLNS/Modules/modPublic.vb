Imports Microsoft.Office.Interop

Module modPublic

    ''' <summary>
    ''' Kiểm tra file Configuration có đủ các thông số hay chưa để yêu cầu người dùng bổ sung thông tin trước khi sử dụng chương trình
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CheckSettings() As Boolean
        Dim bReturn As Boolean = True
        Dim _SQLHelper As New DBAccess
        Dim StrSQL As String = "SELECT * FROM SysVar WHERE ID_DonVi = " & IdDONVI
        Dim dt As DataTable = _SQLHelper.getDataTable(StrSQL)
        If Not (dt.Rows.Count > 0 AndAlso (My_CInt(dt.Rows(0)("MAX_KY"), 0) > 0 And My_CInt(dt.Rows(0)("TINH_THUE_TNCN"), 0) >= 0 And dt.Rows(0)("GIAMDOC") <> "" And dt.Rows(0)("PHOGIAMDOC") <> "" And dt.Rows(0)("TRUONGHCTC") <> "" And dt.Rows(0)("KETOANTRUONG") <> "")) Then
            bReturn = False
        End If
        Return bReturn
    End Function

    ''' <summary>
    ''' Thay cho hàm CStr
    ''' </summary>
    Function My_CStr(InputValue As Object, Optional DefaultValue As String = "") As String
        Dim sReturn As String = ""
        Try
            sReturn = CStr(InputValue)
        Catch ex As Exception
            sReturn = DefaultValue
        End Try
        If sReturn Is Nothing Then sReturn = DefaultValue
        Return sReturn
    End Function

    Function My_CStrByLength(ByVal obj As Object, ByVal iLen As Integer, ByVal cChr As Char) As String
        Dim result As String
        Try
            result = CStr(obj)
        Catch ex As Exception
            result = ""
        End Try
        If Len(result) < iLen Then
            result = StrDup(iLen - Len(result), cChr) & result
        End If
        Return result
    End Function

    ''' <summary>
    ''' Thay cho hàm CStr
    ''' </summary>
    Function My_CBool(InputValue As Object, DefaultValue As String) As Boolean
        Dim bReturn As String
        Try
            bReturn = CBool(InputValue)
        Catch ex As Exception
            bReturn = DefaultValue
        End Try
        Return bReturn
    End Function

    ''' <summary>
    ''' Thay cho hàm CInt
    ''' </summary>
    Function My_CInt(InputValue As Object, DefaultValue As Integer) As Integer
        Dim iReturn As Integer
        Try
            iReturn = IIf(InputValue Is Nothing, DefaultValue, CInt(InputValue))
        Catch ex As Exception
            iReturn = DefaultValue
        End Try
        Return iReturn
    End Function

    ''' <summary>
    ''' Hiển thị nhanh các thông báo, thay cho hàm MessageBox.Show
    ''' </summary>
    Function My_MessageBox(Info As String) As System.Windows.Forms.DialogResult
        Select Case Right(Info, 1)
            Case "."
                Return MessageBox.Show(Info, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Case "!"
                Return MessageBox.Show(Info, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case "?"
                Return MessageBox.Show(Info, "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        End Select
    End Function

    ''' <summary>
    ''' Load dữ liệu vào Combobox cbo, nội dung nằm ở DataTable dt dạng "ID|Ten..."
    ''' ID đưa vào ValueMember còn Tên sẽ đưa vào DisplayMember
    ''' </summary>
    Sub cbo_Binding(ByVal cbo As System.Windows.Forms.ComboBox, ByVal dt As DataTable)
        Try
            With cbo
                .DataSource = Nothing
                .BindingContext = New BindingContext
                .DataSource = dt
                Select Case dt.Columns.Count
                    Case 2 'Có 2 cột: ID và Tên
                        .DisplayMember = dt.Columns(1).Caption
                        .ValueMember = dt.Columns(0).Caption
                    Case 1 'Chỉ có 1 cột tên
                        .DisplayMember = dt.Columns(0).Caption
                End Select
                .SelectedIndex = -1
            End With
            'If cbo.Items.Count > 0 Then cbo.SelectedIndex = 0
        Catch ex As Exception
            If Err.Number <> 91 Then MessageBox.Show(ex.Message, "Error-cbo_LoadData", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Load dữ liệu vào CheckedListBox chkl, nội dung nằm ở DataTable dt dạng "ID|Ten..."
    ''' ID đưa vào ValueMember còn Tên sẽ đưa vào DisplayMember
    ''' </summary>
    Sub chkl_Binding(ByVal chkl As System.Windows.Forms.CheckedListBox, ByVal dt As DataTable)
        Try
            With chkl
                .DataSource = Nothing
                .BindingContext = New BindingContext
                .DataSource = dt
                Select Case dt.Columns.Count
                    Case 2 'Có 2 cột: ID và Tên
                        .DisplayMember = dt.Columns(1).Caption
                        .ValueMember = dt.Columns(0).Caption
                    Case 1 'Chỉ có 1 cột tên
                        .DisplayMember = dt.Columns(0).Caption
                End Select
            End With
            'If cbo.Items.Count > 0 Then cbo.SelectedIndex = 0
        Catch ex As Exception
            If Err.Number <> 91 Then MessageBox.Show(ex.Message, "Error-chkl_LoadData", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Check những item trong CheckedListBox có ID như ListOfId đưa vào
    ''' </summary>
    Sub chkl_SetItemsCheck(ByVal chkl As System.Windows.Forms.CheckedListBox, ListOfID As String, Optional SeparatorChar As Char = ";")
        Try
            Dim arr() As String = ListOfID.Split(SeparatorChar)
            With chkl
                If .Items.Count > 0 Then
                    For i As Integer = 0 To .Items.Count - 1
                        .SetItemChecked(i, False)
                        For Each s As String In arr
                            If s = .Items(i)(0) Then .SetItemChecked(i, True)
                        Next
                    Next
                End If
            End With
        Catch ex As Exception
            If Err.Number <> 91 Then MessageBox.Show(ex.Message, "Error-chkl_LoadData", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Lấy danh sách những ID của dòng được check trong chkl
    ''' </summary>
    Function chkl_GetItemsCheck(ByVal chkl As System.Windows.Forms.CheckedListBox, Optional SeparatorChar As Char = ";") As String
        Dim sReturn As String = ""
        Try
            With chkl
                If .Items.Count > 0 Then
                    For Each item As Object In .CheckedItems
                        sReturn &= item(0) & ";"
                    Next
                    If sReturn.Length > 0 Then sReturn = sReturn.Substring(0, sReturn.Length - 1)
                End If
            End With
        Catch ex As Exception
            If Err.Number <> 91 Then MessageBox.Show(ex.Message, "Error-chkl_LoadData", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return sReturn
    End Function

    Sub Frm_FitSizeToParent(ByVal frm As Form)
        With MainForm
            If .Width < 200 Or .Height < 200 Then Exit Sub
        End With
        With frm
            frm.StartPosition = FormStartPosition.Manual
            .Top = MainForm.Top + 60
            .Left = MainForm.Left + 15
            .Width = MainForm.Width - 30
            .Height = MainForm.Height - 80
        End With
    End Sub

    Public Sub DGV_LoadData(ByRef DGV As DataGridView, ByVal dt As DataTable)
        'Dim r, c As Integer
        DGV.Rows.Clear()
        DGV_Append(DGV, dt)
    End Sub

    Public Sub DGV_Append(ByRef DGV As DataGridView, ByVal dt As DataTable)
        'Dim r, c As Integer
        If dt Is Nothing Then Exit Sub
        Try
            If dt.Rows.Count > 0 Then
                'DGV.Rows.Add(dt.Rows.Count)
                'For r = dt.Rows.Count To 1 Step -1
                '    For c = 0 To dt.Columns.Count - 1
                '        DGV(c, DGV.Rows.Count - r).Value = dt.Rows(dt.Rows.Count - r)(c)
                '        If IsDBNull(DGV(c, DGV.Rows.Count - r)) Then DGV(DGV.Rows.Count - r, c).Value = ""
                '    Next
                'Next
                For Each row As DataRow In dt.Rows
                    DGV.Rows.Add(row.ItemArray)
                Next
            End If
        Catch ex As Exception
            My_MessageBox(ex.Message & "!")
        End Try
    End Sub

    Public Function SaveToWord(doc As Word.Document, Optional sFileName As String = "Word File", Optional bShowResultMsg As Boolean = False) As Boolean
        Dim bReturn As Boolean = False
        'Dim paramSaveFormat As Word.WdSaveFormat = Word.WdSaveFormat.wdFormatDocument
        Try
            Dim sfd As New SaveFileDialog()
            sfd.Filter = "Document File|*.doc|XML Document File|*.docx"
            sfd.Title = "Lưu ra Word"
            sfd.FileName = sFileName
            If sfd.ShowDialog() = DialogResult.OK Then
                'If IO.Path.GetExtension(sfd.FileName).ToLower = "docx" Then paramSaveFormat = Word.WdSaveFormat.wdFormatXML
                doc.SaveAs(sfd.FileName)
                bReturn = True
                If bShowResultMsg Then My_MessageBox("Xuất file Word thành công tại địa chỉ '" & sfd.FileName & "'.")
            End If
        Catch ex As Exception
            bReturn = True
            My_MessageBox(ex.Message & "!")
            Return bReturn
        End Try
    End Function

    Public Function SaveToExcel(ws As Object, Optional sFileName As String = "Excel File", Optional bShowResultMsg As Boolean = False) As Boolean
        Dim bReturn As Boolean = False
        Try
            Dim sfd As New SaveFileDialog()
            sfd.Filter = "Excel File|*.xls|XML Excel File|*.xmlx"
            sfd.Title = "Lưu ra Excel"
            sfd.FileName = sFileName
            If sfd.ShowDialog() = DialogResult.OK Then
                ws.SaveAs(sfd.FileName)
                bReturn = True
                If bShowResultMsg Then My_MessageBox("Xuất Excel thành công tại địa chỉ '" & sfd.FileName & "'.")
            End If
        Catch ex As Exception
            bReturn = False
            My_MessageBox(ex.Message & "!")
        End Try
        Return bReturn
    End Function

#Region "ImageFunction"
    Private Function ThumbnailAbort() As Boolean
        '--> Empty Callback for Thumbnail Generation
        Return True
    End Function

    Public Function My_CImgToThumbnail(ByVal img As Image, ByVal picBox As PictureBox) As Image
        Return img.GetThumbnailImage(picBox.Width, picBox.Height, AddressOf ThumbnailAbort, IntPtr.Zero)
    End Function

    Public Function My_CImgToThumbnail(ByVal img As Image, ByVal iWidth As Integer, ByVal iHeight As Integer) As Image
        Return img.GetThumbnailImage(iWidth, iHeight, AddressOf ThumbnailAbort, IntPtr.Zero)
    End Function

    ''' <summary>
    ''' Đưa giá trị ảnh đọc từ DB ra pixBox
    ''' </summary>
    Function My_LoadImage(picBox As PictureBox, dbImg As Object) As Boolean
        Dim bReturn As Boolean = True
        If Not IsDBNull(dbImg) AndAlso CType(dbImg, Byte()).Length > 1 Then
            Dim mstream As IO.MemoryStream = New IO.MemoryStream(CType(dbImg, Byte()))
            picBox.Image = My_CImgToThumbnail(Image.FromStream(mstream), picBox)
        Else
            picBox.Image = Nothing
            bReturn = False
        End If
        Return bReturn
    End Function

    Function My_CImgToByte(ByVal img As Image) As Byte()
        'Imports System.IO >> dùng ở đây
        Dim result() As Byte = {New Byte}
        Try
            If Not img Is Nothing Then
                Dim stream As New IO.MemoryStream
                img.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp)
                result = stream.ToArray
                stream.Close()
            Else
                result = Nothing
            End If
        Catch ex As Exception
            My_MessageBox("Gặp lỗi lưu ảnh: " & Err.Description & "!")
        End Try
        Return result
    End Function

    'Function tblHoSoCanBo_UpdatePicture(ByVal CanBo_ID As Long, ByVal img As Image) As Boolean
    '    Try
    '        Dim cmd As New System.Data.SqlClient.SqlCommand("UPDATE tblCanBo_HienTai SET [AnhThe] = @img WHERE CanBo_ID = '" & CanBo_ID & "'")
    '        If Not img Is Nothing Then
    '            cmd.Parameters.AddWithValue("@img", System.Data.SqlDbType.Image).Value = My_CImgToByte(img)
    '        Else
    '            cmd.Parameters.AddWithValue("@img", System.Data.SqlTypes.SqlBytes.Null)
    '        End If
    '        cmd.ExecuteNonQuery()
    '        Return True
    '    Catch ex As Exception
    '        My_MessageBox("Đưa ảnh vào dữ liệu: " & ex.Message & ".")
    '        Return False
    '    End Try
    'End Function
#End Region
End Module
