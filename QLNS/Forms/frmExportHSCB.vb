Imports System.IO
Imports Microsoft.VisualBasic.Strings

Public Class frmExportHSCB

    Public IDCB_Moved As String = ""
    Public IDCN_Moved As Integer
    Public TG_Chuyen As Date
    Private dbconn As DBAccess = New DBAccess
    Private MaCB_Moved As String = ""
    Private MaCN_Moved As String = ""
    Private gDir As String = clsCommon.fcnGetValue("TM_Gui_File")
    Private gBackUpFileDir As String = clsCommon.fcnGetValue("TM_Luu_File_Gui")
    Private sFileName As String = ""
    Private arrID_FK As ArrayList = New ArrayList()

    Private Sub frmExportHSCB_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Close()
    End Sub

    Private Sub frmExportHSCB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmExportHSCB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If clsCommon.fcnFormatLocalRegistry() Then
            MessageBox.Show("Thông số hệ thống của Windows vừa được thay đổi. Chương trình sẽ bị đóng." & vbCrLf & "Yêu cầu người sử dụng đăng nhập lại chương trình.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Application.Exit()
        End If
        MaCB_Moved = dbconn.getString("SELECT MaCB FROM HS_Canbo WHERE IdCanbo='" & IDCB_Moved & "'")
        'MaCN_Moved = dbconn.getString("SELECT ma_so FROM ChiNhanh WHERE (id=" & IDCN_Moved & " and id_goc=1) or (id=" & IDCN_Moved & " and id_goc=0) or (id in (SELECT id_goc FROM Chinhanh WHERE id=" & IDCN_Moved & ") and id_goc=1)")
        MaCN_Moved = dbconn.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IDCN_Moved)
        'sFileName = clsCommon.fcnFormatFileName(4, MaCN_Moved, cboMonth.Text, NudYear.Value)
        sFileName = "N4" & Microsoft.VisualBasic.Left(MaCN_Moved, 3) & MaCB_Moved & ".hrm"

        labCB.Text = "Cán bộ: " & dbconn.getString("SELECT HoTen FROM HS_Canbo WHERE IdCanbo='" & IDCB_Moved & "'")
        labCN.Text = "Chuyển hồ sơ tới: " & dbconn.getString("SELECT ten_goi FROM ChiNhanh WHERE (id=" & IDCN_Moved & " and id_goc=1) or (id=" & IDCN_Moved & " and id_goc=0) or (id in (SELECT id_goc FROM Chinhanh WHERE id=" & IDCN_Moved & ") and id_goc=1)")
        HideProgressBar()
    End Sub

    Sub HideProgressBar()
        ProgressBar1.Visible = False
    End Sub

    Sub ShowProgressBar()
        ProgressBar1.Visible = True
    End Sub

    Sub SetProgress(ByVal iVal As Integer)
        ProgressBar1.Value = iVal
        Me.Refresh()
    End Sub

    Private Sub bntExportHS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntExportHS.Click
        Try
            Dim sSourceFile As String = gDir & sFileName
            Dim sZipFile As String = gDir & Microsoft.VisualBasic.Strings.Left(sFileName, sFileName.Length - 4) & ".rar"
            Dim i As Integer = 0
            Dim mDataset As New DataSet
            Dim mTable As New DataTable
            Dim sKeyFieldName As String = ""
            Dim dbconn As DBAccess = New DBAccess

            If Not System.IO.Directory.Exists(gDir) Then
                MessageBox.Show("Không tồn tại thư mục: " & gDir & vbCrLf & " Thao tác bị hủy bỏ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If

            ShowProgressBar()
            ProgressBar1.Maximum = 10
            SetProgress(2)
            If clsCommon.fcnFormatLocalRegistry() Then
                MessageBox.Show("Thông số hệ thống của Windows vừa được thay đổi. Chương trình sẽ bị đóng." & vbCrLf & "Yêu cầu người sử dụng đăng nhập lại chương trình.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Application.Exit()
            End If

            If FileIO.FileSystem.FileExists(sSourceFile) Then
                FileIO.FileSystem.DeleteFile(sSourceFile)
            End If

            If FileIO.FileSystem.FileExists(sZipFile) Then
                FileIO.FileSystem.DeleteFile(sZipFile)
            End If

            ' Lay thong tin cb trong cac Ho so
            mDataset = fcnExecQuery("SELECT * FROM Getdata WHERE cnB =1", "GetData")
            mTable = mDataset.Tables("GetData")
            If mTable.IsInitialized Then
                For i = 0 To mTable.Rows.Count - 1
                    sKeyFieldName = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='" & mTable.Rows(i).Item(1).ToString.Trim & "' and xtype='U') and colid='1'")
                    prcdlCBgui_CN(mTable.Rows(i).Item(1).ToString, MaCB_Moved, sKeyFieldName, arrID_FK)
                Next
            End If
            SetProgress(6)

            ' Nen file va copy vao thu muc luu file gui di
            'clsCommon.ZipIT(sSourceFile, sZipFile)
            Dim tmp As String = FileIO.FileSystem.GetName(sZipFile)
            Dim sBackUpdir As String = ""
            sBackUpdir = gBackUpFileDir & "\" & tmp
            FileIO.FileSystem.CopyFile(sZipFile, sBackUpdir, True)
            FileIO.FileSystem.DeleteFile(sSourceFile)

            ' luu thong tin HS can bo chuyen di
            insertCN_Chuyen_CN(TG_Chuyen, IdDONVI, IDCN_Moved, 1, IDCB_Moved)

            labErr.Text = "File dữ liệu gửi đi được tạo thành công " & vbCr
            labDir.Text = "tại địa chỉ: " & sZipFile
            SetProgress(10)

        Catch ex As Exception
            labErr.Text = ""
            labDir.Text = ""
            MsgBox("Gửi dữ liệu có lỗi " & ex.Message, MsgBoxStyle.Information, "Thông báo")
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, Me.Name, DONVI)
            SetProgress(0)
            HideProgressBar()
            Exit Sub
        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub prcdlCBgui_CN(ByVal sTableName As String, ByVal sMaCB As String, ByVal sKeyFieldName As String, ByRef vArrID_FK As ArrayList)
        Dim sSQL As String = ""
        Dim mDataset As New DataSet
        Dim mTable As New DataTable
        Dim i As Integer = 0
        Dim str As String = ""
        Dim objStreamWriter As StreamWriter
        Dim sDir As String = ""
        Dim sTemp As String = ""
        Dim sKeyField_Name As String = ""

        Try
            Select Case sTableName
                Case "HS_KhenThuong_CT"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_KhenThuong' and xtype='U') and colid='1'")
                    prcdlCBgui_CN("HS_KhenThuong", sMaCB, sKeyField_Name, vArrID_FK)
                    sSQL = "SELECT * FROM HS_KhenThuong_CT Where IdCN_TT='" & IDCB_Moved & "' "
                Case "HS_KhenThuong"
                    sSQL = "SELECT * FROM HS_KhenThuong Where IdKhenThuong in (SELECT distinct IdKhenThuong FROM HS_KhenThuong_CT WHERE IdCN_TT='" & IDCB_Moved & "')"
                Case "HS_NCKH"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='DeTaiNCKH' and xtype='U') and colid='1'")
                    prcdlCBgui_CN("DeTaiNCKH", sMaCB, sKeyField_Name, vArrID_FK)
                    sSQL = "SELECT * FROM HS_NCKH Where IdCanBo='" & IDCB_Moved & "' "
                Case "DeTaiNCKH"
                    sSQL = "SELECT * FROM DeTaiNCKH WHERE IdDeTai in (SELECT distinct IdDeTai FROM HS_NCKH WHERE IdCanBo='" & IDCB_Moved & "')"
                Case "CB_TaiNan"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_VuTaiNan' and xtype='U') and colid='1'")
                    prcdlCBgui_CN("HS_VuTaiNan", sMaCB, sKeyField_Name, vArrID_FK)
                    sSQL = "SELECT * FROM CB_TaiNan Where IdCanBo='" & IDCB_Moved & "' "
                Case "HS_VuTaiNan"
                    sSQL = "SELECT * FROM HS_VuTaiNan WHERE IdVuTaiNan in (SELECT distinct IdVuTN FROM CB_TaiNan WHERE IdCanBo='" & IDCB_Moved & "')"
                Case "HS_DangVien"
                    sSQL = "SELECT * FROM HS_DangVien Where IdCanBo='" & IDCB_Moved & "' "
                Case "HS_Dang"
                    sSQL = "SELECT * FROM HS_Dang WHERE IdDangVien in (SELECT distinct IdDangVien FROM HS_DangVien WHERE IdCanBo='" & IDCB_Moved & "')"
                Case "HS_DoanVien"
                    sSQL = "SELECT * FROM HS_DoanVien Where IdCanBo='" & IDCB_Moved & "' "
                Case "HS_Doan"
                    sSQL = "SELECT * FROM HS_Doan WHERE IdDoanVien in (SELECT distinct IdDoanVien FROM HS_DoanVien WHERE IdCanBo='" & IDCB_Moved & "')"
                Case "HS_CongDoan"
                    sSQL = "SELECT * FROM HS_CongDoan Where IdCanBo='" & IDCB_Moved & "' "
                Case "HS_CongDoanQT"
                    sSQL = "SELECT * FROM HS_CongDoanQT WHERE IdCongDoan in (SELECT distinct IdCongDoan FROM HS_CongDoan WHERE IdCanBo='" & IDCB_Moved & "')"
                Case Else
                    sSQL = "SELECT * FROM " & sTableName & " Where IdCanBo='" & IDCB_Moved & "' "
            End Select

            sDir = gDir & sFileName
            objStreamWriter = File.AppendText(sDir)

            mDataset = fcnExecQuery(sSQL, sTableName)
            mTable = mDataset.Tables(sTableName)
            If mTable.IsInitialized Then
                Dim j As Integer = 0
                Dim IDCanbo_Move As String = ""
                For i = 0 To mTable.Rows.Count - 1
                    With mTable.Rows(i)
                        str = ""
                        For j = 0 To mTable.Columns.Count - 1
                            If UCase(mTable.Columns(j).ColumnName) = UCase(sKeyFieldName) Then
                                str = str & "N'KEY_" & MaCN_Moved & "' ,"
                                If UCase(sTableName) = "HS_CANBO" Then
                                    IDCanbo_Move = .Item(j)
                                ElseIf UCase(sTableName) = "HS_KHENTHUONG" Or UCase(sTableName) = "DETAINCKH" Or UCase(sTableName) = "HS_VUTAINAN" Then
                                    vArrID_FK.Add(.Item(j))
                                End If
                            Else
                                If (UCase(mTable.Columns(j).ColumnName) = "IDCANBO") Or (UCase(mTable.Columns(j).ColumnName) = "IDCN_TT") Then
                                    str = str & "N'KEY_" & MaCN_Moved & "_CB' ,"
                                Else
                                    Select Case UCase(TypeName(.Item(j)))
                                        Case "STRING"
                                            sTemp = Replace(.Item(j), Chr(13) & Chr(10), " - ")
                                            sTemp = Replace(sTemp, Chr(13), " - ")
                                            sTemp = Replace(sTemp, Chr(10), " - ")
                                            sTemp = Replace(sTemp, "|", " ")
                                            Select Case UCase(sTableName)
                                                Case "HS_CANBO"
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDOLD") Then
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & IDCanbo_Move & ";' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If

                                                Case "HS_KHENTHUONG_CT"
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDKHENTHUONG") Then
                                                        Dim k As Integer = 0
                                                        For k = 0 To vArrID_FK.Count - 1
                                                            If vArrID_FK(k) = sTemp Then Exit For
                                                        Next
                                                        str = str & "N'KEY_" & MaCN_Moved & "_FK_" & k & "' ," 'HSKHENTHUONG_IDKHENTHUONG' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_NCKH"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDDETAI") Then
                                                        str = str & "N'KEY_" & MaCN_Moved & "_FK_" & k & "' ," 'DETAINCKH_IDDETAI' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "CB_TAINAN"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDVUTN") Then
                                                        str = str & "N'KEY_" & MaCN_Moved & "_FK_" & k & "' ," 'HSVUTAINAN_IDVUTN' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_DANG"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDDANGVIEN") Then
                                                        str = str & "N'KEY_" & MaCN_Moved & "_FK_" & k & "' ," 'HSDANGVIEN_IDDANGVIEN' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_DOAN"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDDOANVIEN") Then
                                                        str = str & "N'KEY_" & MaCN_Moved & "_FK_" & k & "' ," 'HSDOANVIEN_IDDOANVIEN' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_CONGDOANQT"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDCONGDOAN") Then
                                                        str = str & "N'KEY_" & MaCN_Moved & "_FK_" & k & "' ," 'HSCONGDOAN_IDCONGDOAN' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case Else
                                                    str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                            End Select
                                            'If (UCase(mTable.Columns(j).ColumnName) = "IDOLD") Then
                                            '    str = str & "N'" & Replace(sTemp, "'", "''") & IDCanbo_Move & ";' ,"
                                            'Else
                                            '    str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                            'End If
                                        Case "DATE", "DATETIME"
                                            str = str & "'" & Replace(.Item(j), "'", "''") & "' ,"
                                        Case "DBNULL"
                                            str = str & "'" & .Item(j) & "' ,"
                                        Case "BOOLEAN"
                                            str = str & "" & CInt(.Item(j)) & " ,"
                                        Case "DECIMAL"
                                            str = str & "" & CDbl(.Item(j)) & " ,"
                                        Case "FLOAT", "DOUBLE", "INT", "NUMERIC"
                                            If Double.IsNaN(.Item(j)) Or Double.IsInfinity(.Item(j)) Or Double.IsNegativeInfinity(.Item(j)) Or Double.IsPositiveInfinity(.Item(j)) Then
                                                str = str & " 0,"
                                            Else
                                                str = str & CDbl(.Item(j)) & " ,"
                                            End If
                                            'If Double.IsNaN(.Item(j)) Then
                                            '    str = str & " 0,"
                                            'Else
                                            '    str = str & CDbl(.Item(j)) & " ,"
                                            'End If
                                        Case Else
                                            str = str & "" & .Item(j) & " ,"
                                    End Select
                                End If
                            End If
                        Next
                        str = Microsoft.VisualBasic.Strings.Left(str, Len(str) - 1)
                        str = IdDONVI & "|" & TG_Chuyen & "|" & sTableName & "|" & sKeyFieldName & "|Insert into " & sTableName & " values( " & str & ");"

                        objStreamWriter.WriteLine(str)
                    End With
                Next
            End If

            objStreamWriter.Close()
            mTable = Nothing
            mDataset = Nothing

            Select Case sTableName
                Case "HS_DangVien"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_Dang' and xtype='U') and colid='1'")
                    prcdlCBgui_CN("HS_Dang", sMaCB, sKeyField_Name, vArrID_FK)
                Case "HS_DoanVien"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_Doan' and xtype='U') and colid='1'")
                    prcdlCBgui_CN("HS_Doan", sMaCB, sKeyField_Name, vArrID_FK)
                Case "HS_CongDoan"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_CongDoanQT' and xtype='U') and colid='1'")
                    prcdlCBgui_CN("HS_CongDoanQT", sMaCB, sKeyField_Name, vArrID_FK)
            End Select

        Catch ex As Exception
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, Me.Name, DONVI)
            SetProgress(0)
            HideProgressBar()
            Exit Sub
        End Try

    End Sub

   
End Class