Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic.Strings

Public Class frmChuyenHSCB_CN_ve_TW

    Private gDir As String = clsCommon.fcnGetValue("TM_Gui_File")

    Private Sub frmChuyenHSCB_CN_ve_TW_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmChuyenHSCB_CN_ve_TW_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initGrid()
        bindGrid()
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub bntConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntConvert.Click
        Try
            Dim i As Integer = 0
            Dim mDataset As New DataSet
            Dim mTable As New DataTable
            Dim sKeyFieldName As String = ""
            Dim dbconn As DBAccess = New DBAccess
            Dim dtCB As DataTable
            Dim idxCB As Integer = 0
            Dim _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()

            gDir = "D:\QLNS\Sendfiles\"
            If Not System.IO.Directory.Exists(gDir) Then
                MessageBox.Show("Không tồn tại thư mục: " & gDir & vbCrLf & " Thao tác bị hủy bỏ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If

            ' Lấy danh sách các cán bộ của chi nhánh do TW quản lý
            dtCB = getListCBCN_TWQL()

            If dtCB.Rows.Count > 0 Then
                'If clsCommon.fcnFormatLocalRegistry() Then
                '    MessageBox.Show("Thông số hệ thống của Windows vừa được thay đổi. Chương trình sẽ bị đóng." & vbCrLf & "Yêu cầu người sử dụng đăng nhập lại chương trình.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                '    Application.Exit()
                'End If

                ShowProgressBar()
                ProgressBar1.Maximum = 10 + dtCB.Rows.Count
                SetProgress(2)

                For idxCB = 0 To dtCB.Rows.Count - 1
                    Try
                        Dim bFileName As String = ""
                        Dim bMaCB As String = dtCB.Rows(idxCB)("MaCB").ToString.Trim
                        bFileName = "N5" & bMaCB & ".hrm"
                        Dim bSourceFile As String = gDir & bFileName
                        Dim arrID_FK As ArrayList = New ArrayList()

                        ' Xuất dữ liệu của cán bộ ra file
                        If FileIO.FileSystem.FileExists(bSourceFile) Then FileIO.FileSystem.DeleteFile(bSourceFile)

                        ' Lay thong tin cb trong cac Ho so
                        mDataset = fcnExecQuery("SELECT * FROM Getdata WHERE cnB =1", "GetData")
                        mTable = mDataset.Tables("GetData")
                        If mTable.IsInitialized Then
                            For i = 0 To mTable.Rows.Count - 1
                                sKeyFieldName = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='" & mTable.Rows(i).Item(1).ToString.Trim & "' and xtype='U') and colid='1'")
                                expHS_CBCN_TWQL(mTable.Rows(i).Item(1).ToString, dtCB.Rows(idxCB)("IdCB").ToString(), bMaCB, sKeyFieldName, bSourceFile, arrID_FK)
                            Next
                        End If

                        ' Nếu cán bộ đã được chuyển hồ sơ về TW nhập thì xóa trước khi đọc vào
                        Dim idCBTW As String = ""
                        idCBTW = dbconn.getString("SELECT IdcanBo FROM HS_CanBo WHERE left(Idcanbo,4)='VBSP' And MaCB='" & bMaCB & "'")
                        If Not (idCBTW Is Nothing) Or idCBTW <> "" Then
                            _HS_CanBo.Delete_Human(idCBTW)
                        End If
                        'dbconn.executeSQL("DELETE FROM HS_CanBo WHERE left(Idcanbo,4)='VBSP' And MaCB='" & bMaCB & "'")
                        Threading.Thread.Sleep(100)

                        ' Đọc dữ liệu của cán bộ từ file vào
                        impHS_CBCN_TWQL(bMaCB, bSourceFile)

                        '' Xóa hồ sơ CB do chi nhánh tạo
                        ' voi QLNS tap trung: tap thoi rao lai, vi chua xac dinh nguyen nhan loi doan nay
                        'idCBTW = ""
                        'idCBTW = dbconn.getString("SELECT IdcanBo FROM HS_CanBo WHERE left(Idcanbo,4)<>'VBSP' And MaCB='" & bMaCB & "'")
                        '_HS_CanBo.Delete_Human(idCBTW)
                        ''dbconn.executeSQL("DELETE FROM HS_CanBo WHERE left(Idcanbo,4)<>'VBSP' And MaCB='" & bMaCB & "'")

                        SetProgress(5 + idxCB)

                    Catch ex As Exception
                        MsgBox("Có lỗi" & ex.Message)
                    End Try
                Next

            End If
            
            SetProgress(10 + dtCB.Rows.Count)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub initGrid()

        gridDS_CB.Columns().Clear()
        gridDS_CB.Columns.Add("STT", "STT")
        gridDS_CB.Columns.Add("DonVi", "Đơn vị")
        gridDS_CB.Columns.Add("PhongBan", "Phòng/Ban")
        gridDS_CB.Columns.Add("ChucDanh", "Chức vụ")
        gridDS_CB.Columns.Add("HoTen", "Họ tên")
        gridDS_CB.Columns.Add("MaCB", "Mã cán bộ")
        'Căn chỉnh tiêu đề
        gridDS_CB.Columns("STT").Width = 40
        gridDS_CB.Columns("DonVi").Width = 250 'System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        gridDS_CB.Columns("PhongBan").Width = 220
        gridDS_CB.Columns("ChucDanh").Width = 150
        gridDS_CB.Columns("HoTen").Width = 150
        gridDS_CB.Columns("MaCB").Width = 90
        gridDS_CB.Columns(5).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub bindGrid()
        Try
            Dim dt As DataTable
            Dim j As Integer = 0

            dt = getListCBCN_TWQL()
            gridDS_CB.DataSource = Nothing
            gridDS_CB.Refresh()
            For j = 0 To dt.Rows.Count - 1
                gridDS_CB.Rows.Add()
                gridDS_CB.Rows(j).Cells("STT").Value = j + 1
                gridDS_CB.Rows(j).Cells("DonVi").Value = dt.Rows(j)("Chinhanh").ToString()
                gridDS_CB.Rows(j).Cells("PhongBan").Value = dt.Rows(j)("PhongBan").ToString()
                gridDS_CB.Rows(j).Cells("ChucDanh").Value = dt.Rows(j)("ChucDanh").ToString()
                gridDS_CB.Rows(j).Cells("HoTen").Value = dt.Rows(j)("tenCB").ToString()
                gridDS_CB.Rows(j).Cells("MaCB").Value = dt.Rows(j)("MaCB").ToString()
            Next

        Catch ex As Exception
            MessageBox.Show("Lỗi: không load được danh sách cán bộ chi nhánh TW quản lý:", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function getListCBCN_TWQL() As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("listCB_TWql")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbCBCN_TWQL")
            Return ds.Tables("tbCBCN_TWQL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Private Sub HideProgressBar()
        ProgressBar1.Visible = False
    End Sub

    Private Sub ShowProgressBar()
        ProgressBar1.Visible = True
    End Sub

    Private Sub SetProgress(ByVal iVal As Integer)
        ProgressBar1.Value = iVal
        Me.Refresh()
    End Sub

    Private Sub expHS_CBCN_TWQL(ByVal sTableName As String, ByVal sIdCB As String, ByVal sMaCB As String, ByVal sKeyFieldName As String, ByVal sDesFile As String, ByRef vArrID_FK As ArrayList)
        Dim sSQL As String = ""
        Dim mDataset As New DataSet
        Dim mTable As New DataTable
        Dim i As Integer = 0
        Dim str As String = ""
        Dim objStreamWriter As StreamWriter
        Dim sTemp As String = ""
        Dim sKeyField_Name As String = ""
        Dim dbconn As DBAccess = New DBAccess

        Try
            Select Case sTableName
                Case "HS_KhenThuong_CT"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_KhenThuong' and xtype='U') and colid='1'")
                    expHS_CBCN_TWQL("HS_KhenThuong", sIdCB, sMaCB, sKeyField_Name, sDesFile, vArrID_FK)
                    sSQL = "SELECT * FROM HS_KhenThuong_CT Where IdCN_TT='" & sIdCB & "' "
                Case "HS_KhenThuong"
                    sSQL = "SELECT * FROM HS_KhenThuong Where IdKhenThuong in (SELECT distinct IdKhenThuong FROM HS_KhenThuong_CT WHERE IdCN_TT='" & sIdCB & "')"
                Case "HS_NCKH"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='DeTaiNCKH' and xtype='U') and colid='1'")
                    expHS_CBCN_TWQL("DeTaiNCKH", sIdCB, sMaCB, sKeyField_Name, sDesFile, vArrID_FK)
                    sSQL = "SELECT * FROM HS_NCKH Where IdCanBo='" & sIdCB & "' "
                Case "DeTaiNCKH"
                    sSQL = "SELECT * FROM DeTaiNCKH WHERE IdDeTai in (SELECT distinct IdDeTai FROM HS_NCKH WHERE IdCanBo='" & sIdCB & "')"
                Case "CB_TaiNan"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_VuTaiNan' and xtype='U') and colid='1'")
                    expHS_CBCN_TWQL("HS_VuTaiNan", sIdCB, sMaCB, sKeyField_Name, sDesFile, vArrID_FK)
                    sSQL = "SELECT * FROM CB_TaiNan Where IdCanBo='" & sIdCB & "' "
                Case "HS_VuTaiNan"
                    sSQL = "SELECT * FROM HS_VuTaiNan WHERE IdVuTaiNan in (SELECT distinct IdVuTN FROM CB_TaiNan WHERE IdCanBo='" & sIdCB & "')"
                Case "HS_DangVien"
                    sSQL = "SELECT * FROM HS_DangVien Where IdCanBo='" & sIdCB & "' "
                Case "HS_Dang"
                    sSQL = "SELECT * FROM HS_Dang WHERE IdDangVien in (SELECT distinct IdDangVien FROM HS_DangVien WHERE IdCanBo='" & sIdCB & "')"
                Case "HS_DoanVien"
                    sSQL = "SELECT * FROM HS_DoanVien Where IdCanBo='" & sIdCB & "' "
                Case "HS_Doan"
                    sSQL = "SELECT * FROM HS_Doan WHERE IdDoanVien in (SELECT distinct IdDoanVien FROM HS_DoanVien WHERE IdCanBo='" & sIdCB & "')"
                Case "HS_CongDoan"
                    sSQL = "SELECT * FROM HS_CongDoan Where IdCanBo='" & sIdCB & "' "
                Case "HS_CongDoanQT"
                    sSQL = "SELECT * FROM HS_CongDoanQT WHERE IdCongDoan in (SELECT distinct IdCongDoan FROM HS_CongDoan WHERE IdCanBo='" & sIdCB & "')"
                Case Else
                    sSQL = "SELECT * FROM " & sTableName & " Where IdCanBo='" & sIdCB & "' "
            End Select

            'sDir = gDir & sFileName
            objStreamWriter = File.AppendText(sDesFile)

            mDataset = fcnExecQuery(sSQL, sTableName)
            mTable = mDataset.Tables(sTableName)
            If mTable.IsInitialized Then
                Dim j As Integer = 0
                For i = 0 To mTable.Rows.Count - 1
                    With mTable.Rows(i)
                        str = ""
                        For j = 0 To mTable.Columns.Count - 1
                            If UCase(mTable.Columns(j).ColumnName) = UCase(sKeyFieldName) Then
                                str = str & "N'KEY_" & sMaCB & "' ,"
                                If UCase(sTableName) = "HS_KHENTHUONG" Or UCase(sTableName) = "DETAINCKH" Or UCase(sTableName) = "HS_VUTAINAN" Then
                                    vArrID_FK.Add(.Item(j))
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
                                            Select Case UCase(sTableName)
                                                Case "HS_KHENTHUONG_CT"
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDKHENTHUONG") Then
                                                        Dim k As Integer = 0
                                                        For k = 0 To vArrID_FK.Count - 1
                                                            If vArrID_FK(k) = sTemp Then Exit For
                                                        Next
                                                        str = str & "N'KEY_" & sMaCB & "_FK_" & k & "' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_NCKH"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDDETAI") Then
                                                        str = str & "N'KEY_" & sMaCB & "_FK_" & k & "' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "CB_TAINAN"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDVUTN") Then
                                                        str = str & "N'KEY_" & sMaCB & "_FK_" & k & "' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_DANG"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDDANGVIEN") Then
                                                        str = str & "N'KEY_" & sMaCB & "_FK_" & k & "' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_DOAN"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDDOANVIEN") Then
                                                        str = str & "N'KEY_" & sMaCB & "_FK_" & k & "' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case "HS_CONGDOANQT"
                                                    Dim k As Integer = 0
                                                    For k = 0 To vArrID_FK.Count - 1
                                                        If vArrID_FK(k) = sTemp Then Exit For
                                                    Next
                                                    If (UCase(mTable.Columns(j).ColumnName) = "IDCONGDOAN") Then
                                                        str = str & "N'KEY_" & sMaCB & "_FK_" & k & "' ,"
                                                    Else
                                                        str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                                    End If
                                                Case Else
                                                    str = str & "N'" & Replace(sTemp, "'", "''") & "' ,"
                                            End Select
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
                        str = sTableName & "|" & sKeyFieldName & "|Insert into " & sTableName & " values( " & str & ");"

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
                    expHS_CBCN_TWQL("HS_Dang", sIdCB, sMaCB, sKeyField_Name, sDesFile, vArrID_FK)
                Case "HS_DoanVien"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_Doan' and xtype='U') and colid='1'")
                    expHS_CBCN_TWQL("HS_Doan", sIdCB, sMaCB, sKeyField_Name, sDesFile, vArrID_FK)
                Case "HS_CongDoan"
                    vArrID_FK.Clear()
                    sKeyField_Name = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='HS_CongDoanQT' and xtype='U') and colid='1'")
                    expHS_CBCN_TWQL("HS_CongDoanQT", sIdCB, sMaCB, sKeyField_Name, sDesFile, vArrID_FK)
            End Select

        Catch ex As Exception
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, Me.Name, DONVI)
            SetProgress(0)
            HideProgressBar()
            Exit Sub
        End Try

    End Sub

    Private Sub impHS_CBCN_TWQL(ByVal sMaCB As String, ByVal sSouFile As String)
        Dim fReader As StreamReader
        Dim arrID_FK As ArrayList = New ArrayList()
        Dim tmpTable As String = ""
        Dim idxarr_FK As Integer = 0
        Dim sline As String = ""
        Dim sSQL() As String = {""}
        Dim sKeyValue As String = ""
        Dim sIDCanbo_New As String = ""

        fReader = File.OpenText(sSouFile)
        Do While Not fReader.EndOfStream
            sline = fReader.ReadLine
            sSQL = sline.Split("|")
            sKeyValue = setKeyValueTable(sSQL(0), sSQL(1))

            Select Case UCase(sSQL(0))
                'Case "HS_NCKH"
                '    MsgBox("debug")
                Case "HS_CANBO"
                    sIDCanbo_New = sKeyValue
                Case "HS_KHENTHUONG", "DETAINCKH", "HS_VUTAINAN", "HS_DANGVIEN", "HS_DOANVIEN", "HS_CONGDOAN"
                    If tmpTable <> UCase(sSQL(0)) Then
                        arrID_FK.Clear()
                        idxarr_FK = 0
                        tmpTable = UCase(sSQL(0))
                    End If
                    arrID_FK.Add(sKeyValue)
            End Select
            sSQL(2) = sSQL(2).Replace("N'KEY_" & sMaCB & "'", "N'" & sKeyValue & "'")
            sSQL(2) = sSQL(2).Replace("N'KEY_" & sMaCB & "_CB'", "N'" & sIDCanbo_New & "'")
            If arrID_FK.Count > 0 Then
                'k ban dau =0
                'Insert into HS_NCKH values( N'VBSP00000000001' ,N'KEY_71001_FK_2' ,N'VBSP00000000368' ,0 ,N'' ,'06/07/2011 7:16:59 PM' ,'06/12/2011 10:53:16 AM' );
                If UCase(sSQL(0)) = "HS_NCKH" Then
                    Dim ipre, iaft, iFK As Integer
                    Dim subFK As String
                    subFK = "N'KEY_" & sMaCB & "_FK_"
                    ipre = sSQL(2).IndexOf(subFK)
                    ipre = ipre + Len(subFK)
                    iaft = sSQL(2).IndexOf("'", ipre)
                    iFK = CInt(sSQL(2).Substring(ipre, iaft - ipre))
                    sSQL(2) = sSQL(2).Replace("N'KEY_" & sMaCB & "_FK_" & iFK & "'", "N'" & arrID_FK(iFK) & "'")
                Else
                    sSQL(2) = sSQL(2).Replace("N'KEY_" & sMaCB & "_FK_" & idxarr_FK & "'", "N'" & arrID_FK(idxarr_FK) & "'")
                End If
            End If
            prcHoSo(sSQL(2))
        Loop
        fReader.Close()

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

End Class