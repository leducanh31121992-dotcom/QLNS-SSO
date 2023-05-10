Imports System.IO
Imports Microsoft.VisualBasic.Strings
Imports System.Data.SqlClient
Imports System.Data
Imports Microsoft.Win32
Imports System

Public Class frmReadBigData
    Private vSourceFile As String = "D:\Receivefiles\DuLieuCN\TestBigFile.hrm"
    Private vTablename As String = "DanhMuc"
    Private vBlock As Long = 10000

    Private Sub bntCreateBigData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCreateBigData.Click
        Dim sSQL As String = ""
        Dim mDataset As New DataSet
        Dim mTable As New DataTable
        Dim k As Integer = 0
        Dim j As Integer = 0
        Dim objStreamWriter As StreamWriter
        Dim str As String = ""
        Dim sTemp As String = ""

        Try

            If FileIO.FileSystem.FileExists(vSourceFile) Then FileIO.FileSystem.DeleteFile(vSourceFile)
            objStreamWriter = File.CreateText(vSourceFile)
            objStreamWriter.Close()

            objStreamWriter = File.AppendText(vSourceFile)

            sSQL = "Select * from " & vTablename
            mDataset = fcnExecQuery(sSQL, vTablename)
            mTable = mDataset.Tables(vTablename)
            If mTable.IsInitialized Then
                For k = 0 To mTable.Rows.Count - 1
                    With mTable.Rows(k)
                        Str = ""
                        For j = 0 To mTable.Columns.Count - 1
                            Select Case UCase(TypeName(.Item(j)))
                                Case "STRING", "DATE", "DATETIME"
                                    sTemp = Replace(.Item(j), Chr(13) & Chr(10), " - ")
                                    sTemp = Replace(sTemp, Chr(13), " - ")
                                    sTemp = Replace(sTemp, Chr(10), " - ")
                                    str = str & "N'" & Replace(sTemp, "'", "''") & "' #"
                                Case "DBNULL"
                                    str = str & "'" & .Item(j) & "' #"
                                Case "BOOLEAN"
                                    str = str & "" & CInt(.Item(j)) & " #"
                                Case Else
                                    str = str & "" & .Item(j) & " #"
                            End Select
                        Next
                        Str = Microsoft.VisualBasic.Strings.Left(Str, Len(Str) - 1)
                        str = str.Trim

                        Dim n As Integer
                        For n = 1 To 9000
                            objStreamWriter.WriteLine(str)
                        Next

                        objStreamWriter.WriteLine(Str)
                    End With
                Next
            End If
            objStreamWriter.Close()

            Label1.Text = "Đã tạo xong."

        Catch ex As Exception
            MsgBox("Lấy dữ liệu có lỗi: " & Err.Description, MsgBoxStyle.Information, "Thông báo")
            Exit Sub
        End Try
    End Sub

    Private Sub bntReadBigData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntReadBigData.Click

        Dim sFileName As String = ""
        Dim sFileNameOut As String = ""
        Dim LineCounter As Long = 0
        Dim fReader As StreamReader
        Dim sline As String = ""
        Dim scontent As String = ""

        'doc du lieu
        Label1.Text = Now.ToString
        fReader = File.OpenText(vSourceFile)
        Do While Not fReader.EndOfStream
            LineCounter = LineCounter + 1
            sline = fReader.ReadLine
            scontent &= sline & " "
            If (LineCounter Mod vBlock) = 0 Then
                Dim dbConnection As New SqlConnection(QLNS_CONNSTR)
                Dim cmd As New SqlCommand(scontent, dbConnection)
                Try
                    dbConnection.Open()
                    cmd.ExecuteNonQuery()
                    dbConnection.Close()
                    scontent = ""
                Catch ex As Exception

                End Try
            End If
            'prcExecQueryDUL(sSQL, vIsError, sMaCN)
            'Label1.Text = i.ToString
        Loop
        fReader.Close()
        Label3.Text = LineCounter.ToString
        Label2.Text = Now.ToString
        'MsgBox("Số dòng" & i)
    End Sub

    Private Sub frmReadBigData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    
    Private Sub bntBulkCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntBulkCopy.Click
        Dim dt As New DataTable()
        Dim LineCounter As Long = 0
        Dim fReader As StreamReader
        Dim sline As String = ""
        Dim i As Integer = 0
        Dim item As Object

        Label1.Text = Now.ToString
        Label3.Text = SaveData(vSourceFile, vTablename).ToString
        Label2.Text = Now.ToString

        'Label1.Text = Now.ToString
        'fReader = File.OpenText(vSourceFile)
        'Do While Not fReader.EndOfStream
        '    LineCounter = LineCounter + 1
        '    sline = fReader.ReadLine
        '    Do While sline IsNot Nothing
        '        If (LineCounter Mod vBlock) = 0 Then
        '            Using cn As New SqlConnection(QLNS_CONNSTR)
        '                cn.Open()
        '                Using copy As New SqlBulkCopy(cn)
        '                    copy.ColumnMappings.Add(0, 0)
        '                    copy.ColumnMappings.Add(1, 1)
        '                    copy.ColumnMappings.Add(2, 2)
        '                    copy.ColumnMappings.Add(3, 3)
        '                    copy.ColumnMappings.Add(4, 4)
        '                    copy.DestinationTableName = vTablename
        '                    copy.WriteToServer(dt)
        '                End Using
        '            End Using
        '            dt = Nothing
        '        Else
        '            Dim data() As String = sline.Split("#"c)
        '            If data.Length > 0 Then
        '                If i = 0 Then
        '                    For Each item In data
        '                        dt.Columns.Add(New DataColumn())
        '                    Next item
        '                    i += 1
        '                End If
        '                Dim row As DataRow = dt.NewRow()
        '                row.ItemArray = data
        '                dt.Rows.Add(row)
        '            End If
        '            sline = fReader.ReadLine()
        '        End If
        '    Loop
        'Loop
        'fReader.Close()
        'Label3.Text = LineCounter.ToString
        'Label2.Text = Now.ToString

    End Sub

    Public Function SaveData(ByVal filePath As String, ByVal tableName As String) As Integer
        Dim line As String
        Dim arr As String()
        Dim i, j, count, batchSize As Integer
        batchSize = 100000
        count = 0

        '--Khởi tạo connection đến CSDL và tạo scheme bảng dữ liệu
        Dim oConnection As SqlConnection = OpenDBConnection()
        Dim oCommand As SqlClient.SqlCommand = New SqlCommand(String.Format("SELECT * FROM {0} WHERE 1= 2", tableName), oConnection)
        Dim oAdapter As SqlClient.SqlDataAdapter = New SqlDataAdapter(oCommand)
        Dim oDt As New DataTable(tableName)
        oAdapter.Fill(oDt)

        '--Khởi tạo đối tượng BulkCopy: Mỗi lần lưu batchSize dòng dữ liệu
        Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(oConnection)
        For i = 0 To oDt.Columns.Count - 1
            bulkCopy.ColumnMappings.Add(i, i)
        Next
        bulkCopy.BatchSize = batchSize
        bulkCopy.BulkCopyTimeout = 3000
        bulkCopy.DestinationTableName = tableName

        '--Đọc file và lưu dữ liệu xuống SQL thông qua BulkCopy
        Dim sr As StreamReader = New StreamReader(filePath)
        Do While sr.EndOfStream = False
            line = sr.ReadLine()
            If Not String.IsNullOrEmpty(line) Then
                arr = line.Split(New Char() {"#"c})
                If arr.Length >= oDt.Columns.Count Then
                    Dim row As DataRow = oDt.NewRow()
                    For j = 0 To oDt.Columns.Count - 1
                        row(j) = arr(j)
                    Next
                    oDt.Rows.Add(row)
                    If count Mod batchSize = 0 Then
                        bulkCopy.WriteToServer(oDt)
                        oDt.Clear()
                    End If
                End If
            End If
            count = count + 1
        Loop
        bulkCopy.WriteToServer(oDt)
        sr.Close()
        oConnection.Close()
        Return count
    End Function
End Class