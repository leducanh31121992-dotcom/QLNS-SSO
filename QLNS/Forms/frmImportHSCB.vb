Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Win32
Imports System
Imports System.IO

Public Class frmImportHSCB

    Dim gFileReceived As String = clsCommon.fcnGetValue("TM_Nhan_File")
    Dim gBackUpFileDir As String = clsCommon.fcnGetValue("TM_Luu_File_Nhan")
    Dim dbconn As DBAccess = New DBAccess
    Dim di As DirectoryInfo
    Dim fiHSCB As FileInfo()
    Dim vTG_Chuyen As Date
    Dim vIDCNA As Integer

    Private Sub frmImportHSCB_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Close()
    End Sub

    Private Sub frmImportHSCB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmImportHSCB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If System.IO.Directory.Exists(gFileReceived) Then
            di = New DirectoryInfo(gFileReceived)
            'fiHSCB = di.GetFiles("N4" & DONVI.Substring(0, 3) & "*.rar", SearchOption.AllDirectories)
            fiHSCB = di.GetFiles("N4*.rar", SearchOption.AllDirectories)
            If fiHSCB.Length > 0 Then
                lab1.Text = "Có " & fiHSCB.Length & " hồ sơ cán bộ chuyển tới đang chờ nhập vào hệ thống"
                bntImportHS.Enabled = True
            Else
                lab1.Text = "Không có hồ sơ nào của cán bộ chuyển tới chờ nhập vào hệ thống"
                bntImportHS.Enabled = False
            End If
        Else
            MessageBox.Show("Không tồn tại thư mục: " & gFileReceived & vbCrLf & " Thao tác bị hủy bỏ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
    End Sub

    Private Sub bntImportHS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImportHS.Click
        'Dim dt As DataTable
        'Dim sFile As String = ""
        'Dim sFileName As String = ""
        'Dim sFileNameOut As String = ""
        'Dim fReader As StreamReader
        'Dim sline As String = ""
        'Dim sSQL() As String = {""}
        'Dim i As Integer = 0
        'Dim j As Integer = 0
        'Dim sMaCB As String = ""
        'Dim IdCB_del As String = ""
        'Dim fileArr() As FileInfo
        'Dim labinfo As String = ""
        'Dim ilenPrg As Integer = 0
        'Dim sKeyValue As String = ""
        'Dim sIDCanbo_New As String = ""
        'Dim listCB_New As String = ""

        'fileArr = di.GetFiles
        'ProgressBar1.Maximum = 10 + fileArr.Length * 5
        'ilenPrg = 2
        'SetProgress(ilenPrg)
        'Try
        '    For i = 0 To fileArr.Length - 1
        '        sFileName = fileArr(i).Name
        '        If sFileName.Substring(0, 2) = "N4" Then
        '            labinfo = ""
        '            sIDCanbo_New = ""
        '            vTG_Chuyen = Nothing

        '            sFile = gFileReceived & "\" & sFileName
        '            sMaCB = sFileName.Substring(5, 5)

        '            'Kiem tra HS can bo import vao da co trong he thong chua
        '            'If dbconn.getString("SELECT MaCB FROM HS_Canbo WHERE MaCB='" & sMaCB & "' and idDonvi=" & IdDONVI) <> "" Then
        '            IdCB_del = ""
        '            IdCB_del = dbconn.getString("SELECT IdCanbo FROM HS_canbo WHERE MaCB='" & sMaCB & "' ")
        '            If IdCB_del <> "" Then
        '                labinfo = "Đã tồn tại mã cán bộ " & sMaCB & " trong hệ thống." & vbCr & "Có cập nhật lại hồ sơ của cán bộ này không ?"
        '            Else
        '                labinfo = "Có import hồ sơ cán bộ có mã số " & sMaCB & " vào hệ thống không?"
        '            End If

        '            SetProgress(4)
        '            If MessageBox.Show(labinfo, "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then

        '                'Neu da ton tai ==> xoa
        '                If IdCB_del <> "" Then
        '                    'truoc khi xoa ghi ra file luu thong tin tai thu muc backup
        '                    backupHSCanBo(sMaCB)
        '                    Dim hsCB_exists As clsHS_CanBo = New clsHS_CanBo()
        '                    hsCB_exists.Delete_Human(IdCB_del)
        '                End If

        '                'Import vao he thong

        '                'Luu file nhan vao thu muc luu
        '                Dim sBackUpdir As String = ""
        '                sBackUpdir = gBackUpFileDir & "\" & sFileName
        '                FileIO.FileSystem.CopyFile(sFile, sBackUpdir, True)

        '                'doc du lieu
        '                sFileName = Dir(sFile)
        '                sFileName = gFileReceived & sFileName
        '                clsCommon.UnzipIT(sFileName, gFileReceived, sFileNameOut)

        '                fReader = File.OpenText(sFileNameOut)
        '                Dim arrID_FK As ArrayList = New ArrayList()
        '                Dim tmpTable As String = ""
        '                Dim idxarr_FK As Integer = 0
        '                Do While Not fReader.EndOfStream
        '                    sline = fReader.ReadLine
        '                    sSQL = sline.Split("|")
        '                    sKeyValue = setKeyValueTable(sSQL(2), sSQL(3))
        '                    'If UCase(sSQL(2)) = "HS_CANBO" Then
        '                    '    sIDCanbo_New = sKeyValue
        '                    '    vIDCNA = sSQL(0)
        '                    '    vTG_Chuyen = sSQL(1)
        '                    'End If
        '                    Select Case UCase(sSQL(2))
        '                        Case "HS_CANBO"
        '                            sIDCanbo_New = sKeyValue
        '                            vIDCNA = sSQL(0)
        '                            vTG_Chuyen = sSQL(1)
        '                        Case "HS_KHENTHUONG", "DETAINCKH", "HS_VUTAINAN", "HS_DANGVIEN", "HS_DOANVIEN", "HS_CONGDOAN"
        '                            If tmpTable <> UCase(sSQL(2)) Then
        '                                arrID_FK.Clear()
        '                                idxarr_FK = 0
        '                                tmpTable = UCase(sSQL(2))
        '                            End If
        '                            arrID_FK.Add(sKeyValue)
        '                    End Select
        '                    'sSQL(4) = sSQL(4).Replace("N'KEY_" & DONVI & "'", "N'" & sKeyValue & "'")
        '                    'sSQL(4) = sSQL(4).Replace("N'KEY_" & DONVI & "_CB'", "N'" & sIDCanbo_New & "'")
        '                    sSQL(4) = sSQL(4).Replace("N'KEY_40100" & "'", "N'" & sKeyValue & "'")
        '                    sSQL(4) = sSQL(4).Replace("N'KEY_40100" & "_CB'", "N'" & sIDCanbo_New & "'")
        '                    If arrID_FK.Count > 0 Then
        '                        If UCase(sSQL(2)) = "HS_NCKH" Then
        '                            Dim ipre, iaft, iFK As Integer
        '                            Dim subFK As String
        '                            subFK = "N'KEY_" & DONVI & "_FK_"
        '                            ipre = sSQL(4).IndexOf(subFK)
        '                            ipre = ipre + Len(subFK)
        '                            iaft = sSQL(4).IndexOf("'", ipre)
        '                            iFK = CInt(sSQL(4).Substring(ipre, iaft - ipre))
        '                            sSQL(4) = sSQL(4).Replace("N'KEY_40100" & "_FK_" & iFK & "'", "N'" & arrID_FK(iFK) & "'")
        '                        Else
        '                            sSQL(4) = sSQL(4).Replace("N'KEY_40100" & "_FK_" & idxarr_FK & "'", "N'" & arrID_FK(idxarr_FK) & "'")
        '                        End If
        '                    End If
        '                    prcHoSo(sSQL(4))
        '                Loop
        '                fReader.Close()
        '                FileIO.FileSystem.DeleteFile(sFileName)
        '                FileIO.FileSystem.DeleteFile(sFileNameOut)

        '                ' luu thong tin HS can bo chuyen di
        '                If sIDCanbo_New <> "" Then
        '                    listCB_New &= IIf(listCB_New = "", "'" & sIDCanbo_New & "'", ",'" & sIDCanbo_New & "'")
        '                End If
        '                insertCN_Chuyen_CN(vTG_Chuyen, vIDCNA, IdDONVI, 2, sIDCanbo_New)
        '            Else
        '                IdCB_del = ""
        '            End If
        '        End If
        '        ilenPrg = ilenPrg + 5

        '        SetProgress(ilenPrg)
        '    Next
        '    If listCB_New <> "" Then
        '        dt = dbconn.SelectDBRows("SELECT Hoten, MaCB FROM HS_canbo WHERE IdCanbo in (" & listCB_New & ")")
        '        labinfo = "Import thành công hồ sơ cán bộ: " & vbCr
        '        For j = 0 To dt.Rows.Count - 1
        '            labinfo &= dt.Rows(j).Item("MaCB") & ":" & dt.Rows(j).Item("Hoten") & vbCr
        '        Next
        '        lab1.Text = dt.Rows.Count & " cán bộ import vào hệ thống"
        '        MsgBox(labinfo, MsgBoxStyle.Information, "Thông báo")
        '    Else
        '        lab1.Text = ""
        '    End If
        '    SetProgress(ProgressBar1.Maximum)

        'Catch ex As Exception
        '    WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, Me.Name, DONVI)
        '    SetProgress(0)
        '    HideProgressBar()
        '    Exit Sub
        'End Try

        'Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        'For Each k As Process In pro
        '    k.Kill()
        'Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim dt As DataTable
        Dim sFile As String = ""
        Dim sFileName As String = ""
        Dim sFileNameOut As String = ""
        Dim fReader As StreamReader
        Dim sline As String = ""
        Dim sSQL() As String = {""}
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim sMaCB As String = ""
        Dim IdCB_del As String = ""
        Dim fileArr() As FileInfo
        Dim labinfo As String = ""
        Dim ilenPrg As Integer = 0
        Dim sKeyValue As String = ""
        Dim sIDCanbo_New As String = ""
        Dim listCB_New As String = ""

        fileArr = di.GetFiles
        ProgressBar1.Maximum = 10 + fileArr.Length * 5
        ilenPrg = 2
        SetProgress(ilenPrg)
        Try
            For i = 0 To fileArr.Length - 1
                sFileName = fileArr(i).Name
                If sFileName.Substring(0, 2) = "N4" Then
                    labinfo = ""
                    sIDCanbo_New = ""
                    vTG_Chuyen = Nothing

                    sFile = gFileReceived & "\" & sFileName
                    sMaCB = sFileName.Substring(5, 5)

                    'Kiem tra HS can bo import vao da co trong he thong chua
                    'If dbconn.getString("SELECT MaCB FROM HS_Canbo WHERE MaCB='" & sMaCB & "' and idDonvi=" & IdDONVI) <> "" Then
                    IdCB_del = ""
                    IdCB_del = dbconn.getString("SELECT IdCanbo FROM HS_canbo WHERE MaCB='" & sMaCB & "' ")
                    If IdCB_del <> "" Then
                        labinfo = "Đã tồn tại mã cán bộ " & sMaCB & " trong hệ thống." & vbCr & "Có cập nhật lại hồ sơ của cán bộ này không ?"
                    Else
                        labinfo = "Có import hồ sơ cán bộ có mã số " & sMaCB & " vào hệ thống không?"
                    End If

                    SetProgress(4)
                    If MessageBox.Show(labinfo, "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then

                        'Neu da ton tai ==> xoa
                        If IdCB_del <> "" Then
                            'truoc khi xoa ghi ra file luu thong tin tai thu muc backup
                            backupHSCanBo(sMaCB)
                            Dim hsCB_exists As clsHS_CanBo = New clsHS_CanBo()
                            hsCB_exists.Delete_Human(IdCB_del)
                        End If

                        'Import vao he thong

                        'Luu file nhan vao thu muc luu
                        Dim sBackUpdir As String = ""
                        sBackUpdir = gBackUpFileDir & "\" & sFileName
                        FileIO.FileSystem.CopyFile(sFile, sBackUpdir, True)

                        'doc du lieu
                        sFileName = Dir(sFile)
                        sFileName = gFileReceived & sFileName
                        'clsCommon.UnzipIT(sFileName, gFileReceived, sFileNameOut)

                        fReader = File.OpenText(sFileNameOut)
                        Dim arrID_FK As ArrayList = New ArrayList()
                        Dim tmpTable As String = ""
                        Dim idxarr_FK As Integer = 0
                        Do While Not fReader.EndOfStream
                            sline = fReader.ReadLine
                            sSQL = sline.Split("|")
                            sKeyValue = setKeyValueTable(sSQL(2), sSQL(3))
                            'If UCase(sSQL(2)) = "HS_CANBO" Then
                            '    sIDCanbo_New = sKeyValue
                            '    vIDCNA = sSQL(0)
                            '    vTG_Chuyen = sSQL(1)
                            'End If
                            Select Case UCase(sSQL(2))
                                Case "HS_CANBO"
                                    sIDCanbo_New = sKeyValue
                                    vIDCNA = sSQL(0)
                                    vTG_Chuyen = sSQL(1)
                                Case "HS_KHENTHUONG", "DETAINCKH", "HS_VUTAINAN", "HS_DANGVIEN", "HS_DOANVIEN", "HS_CONGDOAN"
                                    If tmpTable <> UCase(sSQL(2)) Then
                                        arrID_FK.Clear()
                                        idxarr_FK = 0
                                        tmpTable = UCase(sSQL(2))
                                    End If
                                    arrID_FK.Add(sKeyValue)
                            End Select
                            sSQL(4) = sSQL(4).Replace("N'KEY_" & DONVI & "'", "N'" & sKeyValue & "'")
                            sSQL(4) = sSQL(4).Replace("N'KEY_" & DONVI & "_CB'", "N'" & sIDCanbo_New & "'")
                            If arrID_FK.Count > 0 Then
                                If UCase(sSQL(2)) = "HS_NCKH" Then
                                    Dim ipre, iaft, iFK As Integer
                                    Dim subFK As String
                                    subFK = "N'KEY_" & DONVI & "_FK_"
                                    ipre = sSQL(4).IndexOf(subFK)
                                    ipre = ipre + Len(subFK)
                                    iaft = sSQL(4).IndexOf("'", ipre)
                                    iFK = CInt(sSQL(4).Substring(ipre, iaft - ipre))
                                    sSQL(4) = sSQL(4).Replace("N'KEY_" & DONVI & "_FK_" & iFK & "'", "N'" & arrID_FK(iFK) & "'")
                                Else
                                    sSQL(4) = sSQL(4).Replace("N'KEY_" & DONVI & "_FK_" & idxarr_FK & "'", "N'" & arrID_FK(idxarr_FK) & "'")
                                End If
                            End If
                            prcHoSo(sSQL(4))
                        Loop
                        fReader.Close()
                        FileIO.FileSystem.DeleteFile(sFileName)
                        FileIO.FileSystem.DeleteFile(sFileNameOut)

                        ' luu thong tin HS can bo chuyen di
                        If sIDCanbo_New <> "" Then
                            listCB_New &= IIf(listCB_New = "", "'" & sIDCanbo_New & "'", ",'" & sIDCanbo_New & "'")
                        End If
                        insertCN_Chuyen_CN(vTG_Chuyen, vIDCNA, IdDONVI, 2, sIDCanbo_New)
                    Else
                        IdCB_del = ""
                    End If
                End If
                ilenPrg = ilenPrg + 5

                SetProgress(ilenPrg)
            Next
            If listCB_New <> "" Then
                dt = dbconn.SelectDBRows("SELECT Hoten, MaCB FROM HS_canbo WHERE IdCanbo in (" & listCB_New & ")")
                labinfo = "Import thành công hồ sơ cán bộ: " & vbCr
                For j = 0 To dt.Rows.Count - 1
                    labinfo &= dt.Rows(j).Item("MaCB") & ":" & dt.Rows(j).Item("Hoten") & vbCr
                Next
                lab1.Text = dt.Rows.Count & " cán bộ import vào hệ thống"
                MsgBox(labinfo, MsgBoxStyle.Information, "Thông báo")
            Else
                lab1.Text = ""
            End If
            SetProgress(ProgressBar1.Maximum)

        Catch ex As Exception
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, Me.Name, DONVI)
            SetProgress(0)
            HideProgressBar()
            Exit Sub
        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each k As Process In pro
            k.Kill()
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''

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

    Private Sub backupHSCanBo(ByVal vMaCB As String)
        Try
            Dim sFileName As String = "B" & Microsoft.VisualBasic.Left(DONVI, 3) & vMaCB & "_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".hrm"
            Dim sBackUpdir As String = gBackUpFileDir & "\" & sFileName
            Dim sZipFile As String = gBackUpFileDir & Microsoft.VisualBasic.Strings.Left(sFileName, sFileName.Length - 4) & ".rar"
            Dim i As Integer = 0
            Dim mDataset As New DataSet
            Dim mTable As New DataTable
            Dim sKeyFieldName As String = ""
            Dim dbconn As DBAccess = New DBAccess

            'If clsCommon.fcnFormatLocalRegistry() Then
            '    MessageBox.Show("Thông số hệ thống của Windows vừa được thay đổi. Chương trình sẽ bị đóng." & vbCrLf & "Yêu cầu người sử dụng đăng nhập lại chương trình.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            '    Application.Exit()
            'End If

            ' Lay thong tin cb trong cac Ho so
            mDataset = fcnExecQuery("SELECT * FROM Getdata WHERE cnB =1", "GetData")
            mTable = mDataset.Tables("GetData")
            If mTable.IsInitialized Then
                For i = 0 To mTable.Rows.Count - 1
                    sKeyFieldName = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='" & mTable.Rows(i).Item(1).ToString.Trim & "' and xtype='U') and colid='1'")
                    prcdlCB_backup(mTable.Rows(i).Item(1).ToString, vMaCB, sKeyFieldName, sBackUpdir)
                Next
            End If
            SetProgress(6)

            ' Nen file va copy vao thu muc luu file gui di
            'clsCommon.ZipIT(sBackUpdir, sZipFile)
            FileIO.FileSystem.DeleteFile(sBackUpdir)

        Catch ex As Exception
            MsgBox("Sao lưu hồ sơ trước khi cập nhật thông tin cán bộ có lỗi", MsgBoxStyle.Information, "Thông báo")
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, Me.Name, DONVI)
            Exit Sub
        End Try
    End Sub

    Private Sub prcHoSo(ByVal sSQL As String)
        Dim oCommand As SqlClient.SqlCommand
        Dim oTrans As SqlClient.SqlTransaction
        Dim oConnection As SqlConnection
        Try
            oConnection = OpenDBConnection()
            oCommand = oConnection.CreateCommand
            oTrans = oConnection.BeginTransaction(IsolationLevel.ReadCommitted)
            With oCommand
                .Connection = oConnection
                .Transaction = oTrans
                .CommandTimeout = 500
                .CommandType = CommandType.Text
                .CommandText = sSQL
                .ExecuteNonQuery()
            End With
            oTrans.Commit()

            oCommand.Connection.Close()
            oCommand.Dispose()
            oCommand = Nothing
            oConnection.Close()
            oConnection.Dispose()


        Catch ex As Exception
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, "Class", DONVI)
            MsgBox(Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub prcdlCB_backup(ByVal sTableName As String, ByVal sMaCB As String, ByVal sKeyFieldName As String, ByVal sDirFile As String)
        Dim sSQL As String = ""
        Dim mDataset As New DataSet
        Dim mTable As New DataTable
        Dim i As Integer = 0
        Dim str As String = ""
        Dim objStreamWriter As StreamWriter
        Dim sTemp As String = ""
        Dim sKeyField_Name As String = ""
        Dim IDCanBo_backup As String = getCanBo_ID(sMaCB, False)

        Try
            Select Case sTableName
                Case "HS_KhenThuong_CT"
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_KhenThuong' and xtype='U') and colid='1'")
                    prcdlCB_backup("HS_KhenThuong", sMaCB, sKeyField_Name, sDirFile)
                    sSQL = "SELECT * FROM HS_KhenThuong_CT Where IdCN_TT='" & IDCanBo_backup & "' "
                Case "HS_KhenThuong"
                    sSQL = "SELECT * FROM HS_KhenThuong Where IdKhenThuong in (SELECT distinct IdKhenThuong FROM HS_KhenThuong_CT WHERE IdCN_TT='" & IDCanBo_backup & "')"
                Case "HS_NCKH"
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='DeTaiNCKH' and xtype='U') and colid='1'")
                    prcdlCB_backup("DeTaiNCKH", sMaCB, sKeyField_Name, sDirFile)
                    sSQL = "SELECT * FROM HS_NCKH Where IdCanBo='" & IDCanBo_backup & "' "
                Case "DeTaiNCKH"
                    sSQL = "SELECT * FROM DeTaiNCKH WHERE IdDeTai in (SELECT distinct IdDeTai FROM HS_NCKH WHERE IdCanBo='" & IDCanBo_backup & "')"
                Case "CB_TaiNan"
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_VuTaiNan' and xtype='U') and colid='1'")
                    prcdlCB_backup("HS_VuTaiNan", sMaCB, sKeyField_Name, sDirFile)
                    sSQL = "SELECT * FROM CB_TaiNan Where IdCanBo='" & IDCanBo_backup & "' "
                Case "HS_VuTaiNan"
                    sSQL = "SELECT * FROM HS_VuTaiNan WHERE IdVuTaiNan in (SELECT distinct IdVuTN FROM CB_TaiNan WHERE IdCanBo='" & IDCanBo_backup & "')"
                Case "HS_DangVien"
                    sSQL = "SELECT * FROM HS_DangVien Where IdCanBo='" & IDCanBo_backup & "' "
                Case "HS_Dang"
                    sSQL = "SELECT * FROM HS_Dang WHERE IdDangVien in (SELECT distinct IdDangVien FROM HS_DangVien WHERE IdCanBo='" & IDCanBo_backup & "')"
                Case "HS_DoanVien"
                    sSQL = "SELECT * FROM HS_DoanVien Where IdCanBo='" & IDCanBo_backup & "' "
                Case "HS_Doan"
                    sSQL = "SELECT * FROM HS_Doan WHERE IdDoanVien in (SELECT distinct IdDoanVien FROM HS_DoanVien WHERE IdCanBo='" & IDCanBo_backup & "')"
                Case "HS_CongDoan"
                    sSQL = "SELECT * FROM HS_CongDoan Where IdCanBo='" & IDCanBo_backup & "' "
                Case "HS_CongDoanQT"
                    sSQL = "SELECT * FROM HS_CongDoanQT WHERE IdCongDoan in (SELECT distinct IdCongDoan FROM HS_CongDoan WHERE IdCanBo='" & IDCanBo_backup & "')"
                Case Else
                    sSQL = "SELECT * FROM " & sTableName & " Where IdCanBo='" & IDCanBo_backup & "' "
            End Select

            objStreamWriter = File.AppendText(sDirFile)

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
                                str = str & "N'KEY_" & sMaCB & "' ,"
                                If UCase(sTableName) = "HS_CANBO" Then
                                    IDCanbo_Move = .Item(j)
                                End If
                            Else
                                If (UCase(mTable.Columns(j).ColumnName) = "IDCANBO") Or (UCase(mTable.Columns(j).ColumnName) = "IDCN_TT") Then
                                    str = str & "N'KEY_" & sMaCB & "_CB' ,"
                                Else
                                    Select Case UCase(TypeName(.Item(j)))
                                        Case "STRING"
                                            sTemp = Replace(.Item(j), Chr(13) & Chr(10), " - ")
                                            sTemp = Replace(sTemp, Chr(13), " - ")
                                            sTemp = Replace(sTemp, Chr(10), " - ")
                                            sTemp = Replace(sTemp, "|", " ")
                                            If (UCase(mTable.Columns(j).ColumnName) = "IDOLD") Then
                                                str = str & "N'" & Replace(sTemp, "'", "''") & IDCanbo_Move & ";' ,"
                                            Else
                                                str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                            End If
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
                                        Case Else
                                            str = str & "" & .Item(j) & " ,"
                                    End Select
                                End If
                            End If
                        Next
                        str = Microsoft.VisualBasic.Strings.Left(str, Len(str) - 1)
                        str = IdDONVI & "|" & Now & "|" & sTableName & "|" & sKeyFieldName & "|Insert into " & sTableName & " values( " & str & ");"

                        objStreamWriter.WriteLine(str)
                    End With
                Next
            End If

            objStreamWriter.Close()
            mTable = Nothing
            mDataset = Nothing

            Select Case sTableName
                Case "HS_DangVien"
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_Dang' and xtype='U') and colid='1'")
                    prcdlCB_backup("HS_Dang", sMaCB, sKeyField_Name, sDirFile)
                Case "HS_DoanVien"
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_Doan' and xtype='U') and colid='1'")
                    prcdlCB_backup("HS_Doan", sMaCB, sKeyField_Name, sDirFile)
                Case "HS_CongDoan"
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_CongDoanQT' and xtype='U') and colid='1'")
                    prcdlCB_backup("HS_CongDoanQT", sMaCB, sKeyField_Name, sDirFile)
            End Select

        Catch ex As Exception
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, Me.Name, DONVI)
            Exit Sub
        End Try
    End Sub

End Class