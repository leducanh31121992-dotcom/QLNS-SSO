Imports System.Text.RegularExpressions
Imports System.IO
''' <summary>
''' Module chứa các hàm chuẩn hoá 1 chuỗi kí tự
''' Author: Nguyễn Thị Thuỳ Giang
''' Date: 11/8/2007
''' </summary>
''' <remarks></remarks>
Module StringUtil

    Public Function isNumeric(ByVal vValue As String) As Boolean
        Dim _isNumber As Regex = New Regex("^\d+$")
        Dim m As Match = _isNumber.Match(vValue)
        Return m.Success
    End Function

    Public Function N2Text(ByVal v As Object) As String
        If v Is DBNull.Value Or v Is Nothing Then
            Return ""
        Else
            Return v.ToString
        End If
    End Function

    Public Function N2Bool(ByVal v As Object) As Boolean
        If v Is DBNull.Value Or v Is Nothing Then
            Return False
        ElseIf Not CBool(v) Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function N2Number(ByVal v As Object) As Double
        If (v Is Nothing) OrElse (v Is DBNull.Value) Then
            N2Number = 0
        Else
            If Val(v) = 0 Then
                N2Number = 0
            Else
                N2Number = CDbl(v)
            End If
        End If
    End Function

    Public Function setQuote(ByVal input As String) As String
        SetQuote = Replace(input, "'", "''", 1)
    End Function

    ''' <summary>
    ''' Hàm chuẩn hoá một chuỗi kí tự: xoá bỏ các dấu cách thừa
    ''' </summary>
    ''' <param name="strInput">Chuỗi số cần chuẩn hoá</param>
    ''' <returns>Chuỗi đã được chuẩn hoá</returns>
    ''' <remarks></remarks>
    Public Function standardizeString(ByVal strInput As String) As String
        Try
            strInput = strInput.Trim
            strInput = strInput.Substring(0, 1).ToUpper & strInput.Substring(1)
            If strInput = "" Then
                Return strInput
            Else
                strInput = Replace(strInput, "'", "''", 1)
                While strInput.IndexOf("  ") > 0
                    strInput = strInput.Replace("  ", " ")
                End While
                Return strInput
            End If
        Catch ex As Exception
            Return strInput
        End Try
    End Function

    ''' <summary>
    ''' Hàm chuẩn hoá Ten: xoá bỏ các dấu cách thừa, viết hoa đầu mỗi từ
    ''' </summary>
    ''' <param name="strInput">Chuỗi số cần chuẩn hoá</param>
    ''' <returns>Chuỗi đã được chuẩn hoá</returns>
    ''' <remarks></remarks>
    Public Function standardizeName(ByVal strInput As String) As String
        Try
            Dim arrTemp As Char()
            strInput = strInput.ToLower.Trim
            If strInput = "" Then
                Return strInput
            Else
                strInput = Replace(strInput, "'", "''", 1)
                Dim i As Integer
                While strInput.IndexOf("  ") > 0
                    strInput = strInput.Replace("  ", " ")
                End While
                arrTemp = strInput.ToCharArray()
                arrTemp(0) = Char.ToUpper(arrTemp(0))
                For i = 0 To arrTemp.Length - 1
                    If arrTemp(i) = Chr(32) Then arrTemp(i + 1) = Char.ToUpper(arrTemp(i + 1))
                Next
                Return New String(arrTemp)
            End If
        Catch ex As Exception
            Return strInput
        End Try
    End Function

    Public Function formatMoneyinTextbox(ByVal vObjTextBox As TextBox) As TextBox
        Dim Text As String
        Dim selStart, i As Integer
        Dim commaCount_Before As Integer = 0
        Dim commaCount_After As Integer = 0
        If vObjTextBox.Text.Trim = "" Then Return vObjTextBox
        Text = vObjTextBox.Text
        selStart = vObjTextBox.SelectionStart
        commaCount_Before = 0
        commaCount_After = 0
        For i = 0 To Text.Length - 1
            If Text.Substring(i, 1) = "," Then commaCount_Before += 1
        Next
        vObjTextBox.Text = formatMoney(Text)
        Text = vObjTextBox.Text
        For i = 0 To Text.Length - 1
            If Text.Substring(i, 1) = "," Then commaCount_After += 1
        Next
        vObjTextBox.SelectionStart = selStart + (commaCount_After - commaCount_Before)
        Return vObjTextBox
    End Function

    Public Function formatMoney(ByVal vValue As String) As String
        Dim i As Integer
        Dim strReturn As String = ""
        Dim sStringTMP As String = ""
        Dim dotIdx As Integer = 0
        If vValue <> "" Then
            vValue = vValue.Replace(",", "") 'Replace(vValue, ",", "")
            If vValue.Substring(0, 1) <> "-" Then
                vValue = StrReverse(vValue)
                dotIdx = vValue.IndexOf(".")
                For i = 0 To vValue.Length - 1
                    If i <= dotIdx Then
                        strReturn &= vValue(i)
                    Else
                        If (i - dotIdx) > 3 And (i - dotIdx) Mod 3 = 1 Then strReturn &= ","
                        If isNumeric(vValue(i)) Then strReturn &= vValue(i)
                    End If
                Next
            Else
                sStringTMP = vValue.Substring(1)
                sStringTMP = StrReverse(sStringTMP)
                dotIdx = sStringTMP.IndexOf(".")
                For i = 0 To sStringTMP.Length - 1
                    If i <= dotIdx Then
                        strReturn &= sStringTMP(i)
                    Else
                        If (i - dotIdx) > 3 And (i - dotIdx) Mod 3 = 1 Then strReturn &= ","
                        If isNumeric(sStringTMP(i)) Then strReturn &= sStringTMP(i)
                    End If
                Next
                strReturn = strReturn + "-"
            End If
            strReturn = StrReverse(strReturn)
            strReturn = Globals.DeleteChar_FirstAndLast(strReturn, ",")
        End If

        Return strReturn
    End Function
    'vValue.Substring(0,1)
    '(19950000.0).ToString("N",new CultureInfo("en-US")) = 19,950,000.00

    Public Function formatDouble(ByVal vValue As String) As String
        Dim i As Integer
        Dim strReturn As String = ""
        If vValue <> "" Then
            For i = 1 To vValue.Length
                If isNumeric(vValue(i - 1)) Or vValue(i - 1) = "." Then strReturn &= vValue(i - 1)
            Next
        End If
        Return strReturn
    End Function

    Public Function formatLenString(ByVal MaxLen As Integer, ByVal vValString As String) As String
        Dim i As Integer
        Dim strReturn As String = ""
        If vValString.Trim.Length = MaxLen Then
            Return vValString
            Exit Function
        End If
        If vValString.Trim.Length < MaxLen Then
            For i = 1 To MaxLen - vValString.Trim.Length
                strReturn = strReturn & "0"
            Next
            strReturn = strReturn & vValString.Trim
        Else
            strReturn = ""
        End If
        Return strReturn
    End Function

    Public Function MoneyValue(ByVal vValue As String) As String
        Return Replace(vValue, ",", "")
    End Function

    ''' <summary>
    ''' Thủ tục lưu file log
    ''' </summary>
    ''' <param name="action"></param>
    ''' <param name="result"></param>
    ''' <remarks></remarks>
    Public Sub writeLog(ByVal action As String, ByVal result As String)
        Try
            Dim fsOutput As FileStream
            Dim gFileLog As String = clsCommon.fcnGetValue("TM_Backup")
            If Not (System.IO.Directory.Exists(gFileLog)) Then
                System.IO.Directory.CreateDirectory(gFileLog)
            End If
            If Not (System.IO.File.Exists(gFileLog & "log.txt")) Then
                fsOutput = New FileStream(gFileLog & "log.txt", FileMode.Create, FileAccess.Write)
            Else
                fsOutput = New FileStream(gFileLog & "log.txt", FileMode.Append, FileAccess.Write)
            End If
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            Dim content As String = ""
            content = Now() & " | " & action.ToUpper & " | " & result
            srOutput.WriteLine(content.ToString())
            srOutput.Close()
            fsOutput.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        Finally
        End Try
    End Sub

End Module
