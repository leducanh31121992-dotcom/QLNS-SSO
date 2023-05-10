Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Text
Imports System
Imports System.Reflection

Public Class frmCheckDL_DanhMuc


    Dim serverB = "localhost"
    Dim DataBaseB = "QLNS_NgheAn"
    Dim UserB = "vbsp"
    Dim PassB = "asdfgh"
    Dim strconn_DB_B As String = "Data Source=" & serverB & ";Initial Catalog=" & DataBaseB & ";Persist Security Info=True;User ID=" & UserB & ";Password=" & PassB
    Dim dbconn As DBAccess = New DBAccess
    Dim dbconnB As DBAccess = New DBAccess(strconn_DB_B)

    Private Sub frmCheckDL_DanhMuc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        labInfo.Text = "Danh check dữ liệu danh mục của chi nhánh " & DataBaseB
    End Sub

    Private Sub bnt_CheckDM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt_CheckDM.Click
        Dim sSQL As String = ""
        Dim sSQL_tb As String = ""
        Dim mDataset As New DataSet
        Dim mTable As New System.Data.DataTable

        Dim dtCNB As New System.Data.DataTable
        Dim sKeyFieldName As String = ""
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = 0

        sSQL = "Select Ten_Bang from GetData where chi_nhanh = 0"
        mDataset = fcnExecQuery(sSQL, "getData")
        mTable = mDataset.Tables("GetData")
        dtCNB = dbconnB.SelectDBRows(sSQL)
        If dtCNB.Rows.Count = mTable.Rows.Count Then
            If mTable.IsInitialized Then
                For i = 0 To mTable.Rows.Count - 1
                    Dim mDataset_TB As New DataSet
                    Dim mTable_TB As New System.Data.DataTable
                    Dim dtCNB_TB As New System.Data.DataTable

                    sKeyFieldName = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='" & mTable.Rows(i).Item(0).ToString.Trim & "' and xtype='U') and colid='1'")
                    sSQL_tb = "SELECT * FROM " & mTable.Rows(i).Item("Ten_Bang") & "  order by " & sKeyFieldName
                    mDataset_TB = fcnExecQuery(sSQL_tb, "tbDetail_1")
                    mTable_TB = mDataset_TB.Tables("tbDetail_1")
                    dtCNB_TB = dbconnB.SelectDBRows(sSQL_tb)
                    If dtCNB_TB.Rows.Count = mTable_TB.Rows.Count Then
                        If mTable_TB.IsInitialized Then
                            For j = 0 To mTable_TB.Rows.Count - 1
                                If Not (mTable_TB.Rows(j).Item(0) = dtCNB_TB.Rows(j).Item(0) And mTable_TB.Rows(j).Item(1) = dtCNB_TB.Rows(j).Item(1) And mTable_TB.Rows(j).Item(2) = dtCNB_TB.Rows(j).Item(2)) Then
                                    WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number & mTable.Rows(j).Item("Ten_Bang") & ":" & dtCNB_TB.Rows(j).Item(0), Me.Name, DONVI)
                                End If
                                'sSQL_tb = "SELECT * FROM " & mTable_TB.Rows(j).Item("Ten_Bang")
                                'mDataset_TB = fcnExecQuery(sSQL, "getData")
                                'mTable_TB = mDataset.Tables("GetData")
                                'dtCNB_TB = dbconnB.SelectDBRows(sSQL)

                            Next
                        End If
                    Else
                        If dtCNB_TB.Rows.Count < mTable_TB.Rows.Count Then
                            MsgBox("Bang " & mTable.Rows(i).Item("Ten_Bang") & " cua Chi nhanh B THIEU danh muc")
                        Else
                            MsgBox("Bang " & mTable.Rows(i).Item("Ten_Bang") & " cua Chi nhanh B THUA danh muc")
                        End If
                    End If
                Next
            End If
            MsgBox("Chi nhanh B danh muc chuẩn")
        Else
            If dtCNB.Rows.Count < mTable.Rows.Count Then
                MsgBox("Chi nhanh B THIEU bang danh muc")
            Else
                MsgBox("Chi nhanh B THUA bang danh muc")
            End If
        End If

    End Sub

    Private Sub cmd_INS_D_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_INS_D.Click
        ImportExcel_DM_Intellect(txtPath.Text)
    End Sub

    Private Sub ImportExcel_DM_Intellect(ByVal PathExcelFile As String)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim db As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim sql As String = ""
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(1).Select()

                Dim vlues As String = ""
                row = 2
                While .cells(row, "A").value.ToString.Trim <> ""
                    If .cells(row, "B").value Is Nothing Then
                        vlues = ""
                    Else
                        vlues = .cells(row, "B").value.ToString.Trim
                    End If

                    sql = "INSERT INTO DM_Intellect (code,descript) values ('" & .cells(row, "A").value.ToString.Trim & "', N'" & vlues & "')"
                    db.executeSQL(sql)
                    row = row + 1
                End While
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next
        MessageBox.Show("OK")

    End Sub

    Private Sub cmdChonTM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChonTM.Click
        Dim openFile As New OpenFileDialog
        Try
            openFile.Filter = "Excel Files (*.xls)|*.xls"
            openFile.ShowDialog()
            txtPath.Text = openFile.FileName
        Catch ex As Exception

        End Try
    End Sub
End Class