Imports System.IO
'Imports ICSharpCode.SharpZipLib.Checksums
'Imports ICSharpCode.SharpZipLib.Zip
'Imports ICSharpCode.SharpZipLib.GZip
Imports Microsoft.Win32
Imports System.Security.Permissions

Public Class clsCommon

    '    Public Shared Sub ZipIT(ByVal sFileNameIn As String, ByVal sFileNameOut As String)


    '        Dim objCrc32 As New Crc32()
    '        Dim strmZipOutputStream As ZipOutputStream

    '        strmZipOutputStream = New ZipOutputStream(File.Create(sFileNameOut))
    '        strmZipOutputStream.SetLevel(6)

    '        REM Compression Level: 0-9
    '        REM 0: no(Compression)
    '        REM 9: maximum compression
    '        Dim sFileName As String = FileIO.FileSystem.GetName(sFileNameIn)
    '        Dim strmFile As FileStream = File.OpenRead(sFileNameIn)
    '        Dim abyBuffer(strmFile.Length - 1) As Byte

    '        strmFile.Read(abyBuffer, 0, abyBuffer.Length)
    '        Dim objZipEntry As ZipEntry = New ZipEntry(sFileName)

    '        objZipEntry.DateTime = DateTime.Now
    '        objZipEntry.Size = strmFile.Length
    '        strmFile.Close()
    '        objCrc32.Reset()
    '        objCrc32.Update(abyBuffer)
    '        objZipEntry.Crc = objCrc32.Value
    '        strmZipOutputStream.PutNextEntry(objZipEntry)
    '        strmZipOutputStream.Write(abyBuffer, 0, abyBuffer.Length)

    '        strmZipOutputStream.Finish()
    '        strmZipOutputStream.Close()

    '    End Sub

    '    Public Shared Sub UnzipIT(ByVal sFileNameIn As String, ByVal sDir As String, ByRef sFileNameOut As String)
    '        Dim s As New ZipInputStream(File.OpenRead(sFileNameIn))
    '        Dim theEntry As ZipEntry
    '        theEntry = s.GetNextEntry()
    '        Do While Not (IsNothing(theEntry))
    '            Dim fileName As String = Path.GetFileName(theEntry.Name)
    '            Directory.CreateDirectory(sDir)
    '            Dim sWriter As FileStream
    '            sWriter = File.Create(sDir + fileName)
    '            sFileNameOut = sDir + fileName
    '            Dim size As Integer = 2048
    '            Dim data(2048) As Byte
    '            Do While (True)
    '                If theEntry.Size = 0 Then Exit Sub
    '                size = s.Read(data, 0, data.Length)
    '                If (size > 0) Then
    '                    sWriter.Write(data, 0, size)
    '                Else
    '                    GoTo ex1
    '                End If
    '            Loop
    'EX1:
    '            sWriter.Close()
    '            theEntry = s.GetNextEntry
    '        Loop
    '        s.Close()
    '    End Sub

    '    Public Shared Function checkEmptyFileZip(ByVal sFileNameIn As String, ByVal sDir As String, ByRef sFileNameOut As String) As Boolean
    '        Dim s As New ZipInputStream(File.OpenRead(sFileNameIn))
    '        Dim theEntry As ZipEntry
    '        Dim _return As Boolean
    '        theEntry = s.GetNextEntry()
    '        Do While Not (IsNothing(theEntry))
    '            If theEntry.Size = 0 Then
    '                _return = True
    '            Else
    '                _return = False
    '            End If
    '            theEntry = s.GetNextEntry
    '        Loop
    '        s.Close()
    '        Return _return
    '    End Function

    Public Shared Function fcnFormatLocalRegistry() As Boolean
        Dim sFormat As String = ""
        Dim fResult As Boolean = False
        'sFormat = fcnReadReg(Server & "\HKEY_CURRENT_USER\Control Panel\International", "sDecimal")
        'If sFormat = "," Then prcWriteReg(Server & "\HKEY_CURRENT_USER\Control Panel\International", "sDecimal", ".")

        sFormat = fcnReadString("\International", "sShortDate")
        'If (sFormat = "dd/MM/yyyy" Or sFormat = "dd/MM/yy") Then
        If (sFormat <> "MM/dd/yyyy") Then
            prcWriteString("\International", "sShortDate", "MM/dd/yyyy")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "sLongDate")
        'If (sFormat = "dd/MM/yyyy" Or sFormat = "dd/MM/yy") Then
        If (sFormat <> "dddd, MMMM dd, yyyy") Then
            prcWriteString("\International", "sLongDate", "dddd, MMMM dd, yyyy")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "sDecimal")
        If (sFormat <> ".") Then
            prcWriteString("\International", "sDecimal", ".")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "sThousand")
        If (sFormat <> ",") Then
            prcWriteString("\International", "sThousand", ",")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "sNegativeSign")
        If (sFormat <> "-") Then
            prcWriteString("\International", "sNegativeSign", "-")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "iNegNumber")
        If (sFormat <> "1") Then
            prcWriteString("\International", "iNegNumber", "1")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "iLZero")
        If (sFormat <> "1") Then
            prcWriteString("\International", "iLZero", "1")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "s1159")
        If (sFormat <> "AM") Then
            prcWriteString("\International", "s1159", "AM")
            fResult = True
        End If
        sFormat = fcnReadString("\International", "s2359")
        If (sFormat <> "PM") Then
            prcWriteString("\International", "s2359", "PM")
            fResult = True
        End If
        fcnFormatLocalRegistry = fResult

    End Function
    ''' <summary>
    ''' sKindFile = 0: File du lieu phat sinh trong thang tinh gui len TW ;
    ''' sKindFile = 1: File phuc hoi du lieu Danh muc; 
    ''' sKindFile = 2: File phuc hoi du lieu Phat sinh cho tinh;
    ''' sKindFile = 3: File Phuc hoi du lieu danh muc va du lieu phat sinh cho tinh
    ''' sKindFile = 4: File du lieu chuyen ngang HS_Canbo giu cac chi nhanh
    ''' </summary>
    ''' <param name="sKindFile"></param>
    ''' <param name="sCode"></param>
    ''' <param name="sMonth"></param>
    ''' <param name="sYear"></param>
    ''' <remarks></remarks>
    Public Shared Function fcnFormatFileName(ByVal sKindFile As Int16, ByVal sCode As String, Optional ByVal sMonth As String = "", Optional ByVal sYear As String = "") As String
        Dim sFormatFileName As String = ""
        Dim vMonth As String = ""
        Select Case sMonth.Trim
            Case "10"
                vMonth = "A"
            Case "11"
                vMonth = "B"
            Case "12"
                vMonth = "C"
            Case Else
                vMonth = sMonth.Trim
        End Select
        sFormatFileName = "N" & sKindFile & Microsoft.VisualBasic.Left(sCode, 3) & vMonth & Microsoft.VisualBasic.Right(sYear, 2) & ".hrm"
        fcnFormatFileName = sFormatFileName
    End Function

    Public Shared Function fcnReadString(ByVal sSection As String, ByVal sIdent As String) As String
        ' Ham doc gia tri cua mot Key tu Registry
        ' sSection la duong dan tinh tu sau HKEY_CURRENT_USER\
        ' sIndent la ten cua Key can lay gia tri
        ' sDefault la gia tri mac dinh tra lai cua ham trong truong hop ko thay Key do

        'Dim oReg As Registry
        Dim oRegKey As RegistryKey
        Dim sTemp As String
        Dim sResult As String = ""
        Dim sDefault As String = ""
        Dim sRoot As String = "Control Panel\"
        Dim sDelimiter As String = ""

        sTemp = sRoot
        If Trim(sSection) <> "" Then
            sTemp = IIf(Right(sTemp, 1) = sDelimiter, sTemp, sTemp & sDelimiter)
            sTemp = sTemp & sSection
        End If

        'oRegKey = oReg.CurrentUser.OpenSubKey(sTemp, False)
        oRegKey = Registry.CurrentUser.OpenSubKey(sTemp, False)
        sResult = oRegKey.GetValue(sIdent, sDefault)
        oRegKey.Close()

        fcnReadString = sResult

    End Function

    Public Shared Sub prcWriteString(ByVal sSection As String, ByVal sIdent As String, ByVal sValue As String)
        ' Ham ghi gia tri cua mot Key vao Registry
        ' sSection la duong dan tinh tu sau HKEY_CURRENT_USER\SOFTWARE\
        ' sIndent la ten cua Key can set gia tri
        ' sValue la gia tri can set cho Key do

        'Dim oReg As Registry
        Dim oRegKey As RegistryKey
        Dim sTemp As String
        Dim sResult As String = ""
        Dim sRoot As String = "Control Panel\"
        Dim sDelimiter As String = ""

        sTemp = sRoot
        If Trim(sSection) <> "" Then
            sTemp = IIf(Right(sTemp, 1) = sDelimiter, sTemp, sTemp & sDelimiter)
            sTemp = sTemp & sSection
        End If

        oRegKey = Registry.CurrentUser.OpenSubKey(sTemp, True)
        oRegKey.SetValue(sIdent, sValue)
        oRegKey.Close()

        'oReRegKey = Registry.CurrentUser.OpenSubKey(sRoot, True)
        'oReRegKey.Close()
        'oReRegKey = Registry.CurrentUser.OpenSubKey(sTemp, True)
        'oReRegKey.Close()
    End Sub

    Public Shared Function fcnReadReg(ByVal sSection As String, ByVal sIdent As String) As String
        ' Ham doc gia tri cua mot Key tu Registry
        ' sSection la duong dan tinh tu sau HKEY_CURRENT_USER\
        ' sIndent la ten cua Key can lay gia tri
        ' sDefault la gia tri mac dinh tra lai cua ham trong truong hop ko thay Key do

        'Dim oReg As Registry
        Dim oRegKey As RegistryKey
        Dim sResult As String = ""
        Dim sDefault As String = ""
        Dim sDelimiter As String = ""

        'sTemp = sRoot
        'If Trim(sSection) <> "" Then
        '    sTemp = IIf(Right(sTemp, 1) = sDelimiter, sTemp, sTemp & sDelimiter)
        '    sTemp = sTemp & sSection
        'End If

        oRegKey = Registry.CurrentUser.OpenSubKey(sSection, False)
        sResult = oRegKey.GetValue(sIdent, sDefault)
        oRegKey.Close()

        fcnReadReg = sResult

    End Function

    Public Shared Sub prcWriteReg(ByVal sSection As String, ByVal sIdent As String, ByVal sValue As String)
        ' Ham ghi gia tri cua mot Key vao Registry
        ' sSection la duong dan tinh tu sau HKEY_CURRENT_USER\SOFTWARE\
        ' sIndent la ten cua Key can set gia tri
        ' sValue la gia tri can set cho Key do

        Dim oRegKey As RegistryKey
        'Dim sTemp As String
        Dim sResult As String = ""

        Dim sDelimiter As String = ""

        'sTemp = sRoot
        'If Trim(sSection) <> "" Then
        '    sTemp = IIf(Right(sTemp, 1) = sDelimiter, sTemp, sTemp & sDelimiter)
        '    sTemp = sTemp & sSection
        'End If

        oRegKey = Registry.CurrentUser.OpenSubKey(sSection, True)
        oRegKey.SetValue(sIdent, sValue)
        oRegKey.Close()

    End Sub

    Public Shared Sub prcLoadListView(ByVal lvwName As Windows.Forms.ListView, ByVal sSQL As String, ByVal sTableName As String, ByVal intNumberOfField As Integer)
        ' Load du lieu len ListView
        Dim myDS As New DataSet
        Dim i As Integer = 0
        Dim j As Integer
        Dim dTable As DataTable
        Dim lvi As Windows.Forms.ListViewItem
        Dim drow As DataRow

        Try
            ' Load du lieu vao DataSet
            myDS = fcnExecQuery(sSQL, sTableName)
            ' Get DataTable
            dTable = myDS.Tables(0)

            With lvwName
                .Items.Clear()

                'Display items in the ListView control
                For i = 0 To dTable.Rows.Count - 1
                    drow = dTable.Rows(i)

                    ' Only row that have not been deleted
                    If (drow.RowState <> DataRowState.Deleted) Then

                        'Define the list items                    
                        lvi = New Windows.Forms.ListViewItem(drow(0).ToString())
                        For j = 1 To intNumberOfField - 1
                            lvi.SubItems.Add(drow(j).ToString())
                        Next j

                        'Add the list items to the ListView
                        lvwName.Items.Add(lvi)
                    End If
                Next i

            End With

            dTable = Nothing
            myDS = Nothing

        Catch ex As Exception
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, "Errors", DONVI)
            Exit Sub
        End Try
    End Sub

    Public Shared Function fcnGetValue(ByVal sGiatri As String) As String
        Dim sReturn$ = ""
        Try
            Dim sSQL As String = "Select " & sGiatri & " from sysVar where ID_DonVi = " & IdDONVI
            Dim conn As New DBAccess
            sReturn = conn.getString(sSQL)
            sReturn = My_CStr(sReturn, "")
        Catch ex As Exception
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, "Errors", DONVI)
        End Try
        Return sReturn
    End Function

    Public Shared Function prcDeleteDB(ByVal vMaCN As String) As Boolean
        Dim vIsError As Boolean = False
        Try
            Dim sSQL() As String = {""}
            Dim strSQL As String = ""
            Dim mDataset As New DataSet
            Dim mTable As New DataTable
            Dim sKeyFieldName As String = ""
            Dim dbconn As DBAccess = New DBAccess
            Dim TenVTCN As String = ""
            Dim i As Integer

            strSQL = "Select Ten_Bang from getdata where chi_nhanh = 1 Order by id desc"
            mDataset = fcnExecQuery(strSQL, "Getdata")
            mTable = mDataset.Tables("getdata")
            TenVTCN = fncShortNameDV(vMaCN)
            If mTable.Rows.Count Then
                For i = 0 To mTable.Rows.Count - 1
                    If UCase(mTable.Rows(i).Item("Ten_bang")) <> "CN_HANGDN" Then
                        sKeyFieldName = dbconn.getString("SELECT [name] FROM syscolumns WHERE [id]=(SELECT [id] FROM sysobjects WHERE [name]='" & mTable.Rows(i).Item("Ten_bang") & "' and xtype='U') and colid='1'")
                        sSQL(0) = "Delete from " & mTable.Rows(i).Item("Ten_bang") & " where left(" & sKeyFieldName & ",4)='" & TenVTCN & "'"
                        'sSQL(0) = "Delete from " & mTable.Rows(i).Item("Ten_bang") & " where left(" & sKeyFieldName & ",4)='TBNA' OR left(" & sKeyFieldName & ",4)='TVPA'"
                        prcExecQueryDUL(sSQL, vIsError, DONVI)
                        If vIsError Then Exit For
                    End If
                Next
            End If
            mTable = Nothing
            mDataset = Nothing
        Catch ex As Exception
            vIsError = True
            WriteToLogFile(My.Application.Info.DirectoryPath + "\Errors\ErrorLog", Err.Description, Err.Number, "Class", vMaCN)
        End Try

        If vIsError Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Shared Function fncShortNameDV(ByVal vMasoDonVi As String) As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            Return dbconn.getString("SELECT ten_vt FROM CHINHANH WHERE ma_so='" & vMasoDonVi & "'")
        Catch ex As Exception
            Return ""
        End Try
    End Function

End Class