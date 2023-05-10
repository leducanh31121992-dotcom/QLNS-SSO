
Public Class DateTimeUtil
    Public Shared Function MinSqlDateTime() As DateTime
        Return New DateTime(1900, 1, 1, 0, 0, 0, 0)
    End Function


    ''' <summary>
    ''' Hàm convert 1 chuỗi ngày tháng dạng 'dd/mm/yyyy' về kiểu Datetime dang 'mm/dd/yyyy'
    ''' </summary>
    ''' <param name="VNDatetime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getDate(ByVal VNDatetime As String) As Date
        Dim temp() As String
        Dim sep() As Char = {"/", ".", "-"}
        Try
            temp = VNDatetime.Split(sep)
            If Convert.ToInt16(temp(1)) And isNumeric(0) And isNumeric(2) Then
                If Convert.ToInt16(temp(0)) <= 31 And Convert.ToInt16(temp(1)) <= 12 And Convert.ToInt16(temp(2)) Then
                    Return Convert.ToDateTime(temp(1) & "/" & temp(0) & "/" & temp(2))
                Else
                    Return New Date
                End If
            Else
                Return New Date
            End If
        Catch ex As Exception
            Return New Date
        End Try
    End Function

    ''' <summary>
    ''' Hàm convert 1 chuỗi ngày tháng dạng 'dd/mm/yyyy hh:mm:ss' về kiểu Datetime dang 'mm/dd/yyyy hh:mm:ss'
    ''' </summary>
    ''' <param name="VNDatetime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getDatetime(ByVal VNDatetime As String) As Date
        Dim temp() As String
        Dim sep() As Char = {"/", ".", "-", " ", ":"}
        Try
            temp = VNDatetime.Split(sep)
            If isNumeric(temp(1)) And isNumeric(temp(0)) And isNumeric(temp(2)) Then
                If Convert.ToInt16(temp(0)) <= 31 And Convert.ToInt16(temp(1)) <= 12 And Convert.ToInt16(temp(2)) <= 9999 And Convert.ToInt16(temp(3)) <= 23 And Convert.ToInt16(temp(3)) >= 0 And Convert.ToInt16(temp(4)) <= 59 And Convert.ToInt16(temp(4)) >= 0 And Convert.ToInt16(temp(5)) >= 0 And Convert.ToInt16(temp(5)) <= 59 Then
                    Return Convert.ToDateTime(temp(1) & "/" & temp(0) & "/" & temp(2) & " " & temp(3) & ":" & temp(4) & ":" & temp(5))
                Else
                    Return New Date
                End If
            Else
                Return New Date
            End If
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' Hàm convert 1 chuỗi ngày tháng dạng 'dd/mm/yyyy' về kiểu Datetime dang 'mm/dd/yyyy hh:mm:ss'
    ''' trong do hh, mm, ss lấy theo thời gian hiện tại
    ''' </summary>
    ''' <param name="VNDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getDateCurrTime(ByVal VNDate As String) As Date
        Dim temp() As String
        Dim sep() As Char = {"/", ".", "-", " ", ":"}
        Try
            temp = VNDate.Split(sep)
            If isNumeric(temp(1)) And isNumeric(temp(0)) And isNumeric(temp(2)) Then
                If Convert.ToInt16(temp(0)) <= 31 And Convert.ToInt16(temp(1)) <= 12 And Convert.ToInt16(temp(2)) <= 9999 Then
                    Return Convert.ToDateTime(temp(1) & "/" & temp(0) & "/" & temp(2) & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second)
                Else
                    Return New Date
                End If
            Else
                Return New Date
            End If
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' Hàm này trả về kiểu string của giờ phút và ngày tháng năm dạng đầy đủ, ví dụ:  dd/mm/yyyy hh:mm:ss
    ''' </summary>
    ''' <param name="dateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getLongDateTimeNoneDay(ByVal dateTime As Date) As String
        Try
            Dim day As String
            day = dateTime.ToString("dd/MM/yyyy") & " " & dateTime.ToString("HH:mm:ss")
            Return day
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Hàm này trả về kiểu string của giờ phút và ngày tháng năm dạng đầy đủ, ví dụ: hh:mm dd/mm/yyyy
    ''' </summary>
    ''' <param name="dateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getLongTimeDateVN(ByVal dateTime As Date) As String
        Try
            Dim day As String
            day = dateTime.ToString("HH:mm") & " " & dateTime.ToString("dd/MM/yyyy")
            Return day
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Hàm này trả về kiểu string của giờ phút, ví dụ: hh:mm
    ''' </summary>
    ''' <param name="dateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getTime(ByVal dateTime As Date) As String
        Try
            Dim time As String
            time = dateTime.ToString("HH:mm")
            Return time
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Hàm này trả về kiểu string của ngày tháng năm dạng ngắn: dd/MM/yyyy
    ''' </summary>
    ''' <param name="dateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 'Shared
    Public Shared Function getShortDate(ByVal dateTime As Date) As String
        Try
            Return dateTime.ToString("dd/MM/yyyy")
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Hàm này trả về kiểu string của ngày tháng năm dạng đầy đủ, ví dụ: Thứ Sáu, 14/04/2006 14:15 PM
    ''' </summary>
    ''' <param name="dateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getLongDateTime(ByVal dateTime As Date) As String
        Try
            Dim day As String
            Select Case (Convert.ToString(dateTime.DayOfWeek))
                Case "Sunday"
                    day = "Chủ Nhật"
                Case "Monday"
                    day = "Thứ Hai"
                Case "Tuesday"
                    day = "Thứ Ba"
                Case "Wednesday"
                    day = "Thứ Tư"
                Case "Thursday"
                    day = "Thứ Năm"
                Case "Friday"
                    day = "Thứ Sáu"
                Case "Saturday"
                    day = "Thứ Bảy"
                Case Else
                    day = ""
            End Select
            day = day & ", " & dateTime.ToString("dd/MM/yyyy - HH:mm") '// Khong co PM/AM
            Return day
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Hàm này trả về kiểu string của ngày tháng năm dạng đầy đủ, ví dụ: Thứ Sáu, 10/08/2008
    ''' </summary>
    ''' <param name="dateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getLongDate(ByVal dateTime As Date) As String
        Try
            Dim day As String
            Select Case (Convert.ToString(dateTime.DayOfWeek))
                Case "Sunday"
                    day = "Chủ Nhật"
                Case "Monday"
                    day = "Thứ Hai"
                Case "Tuesday"
                    day = "Thứ Ba"
                Case "Wednesday"
                    day = "Thứ Tư"
                Case "Thursday"
                    day = "Thứ Năm"
                Case "Friday"
                    day = "Thứ Sáu"
                Case "Saturday"
                    day = "Thứ Bảy"
                Case Else
                    day = ""
            End Select
            day = day & ", " & dateTime.ToString("dd/MM/yyyy") 'Khong co PM/AM
            Return day
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Hàm này trả về kiểu string của giờ phút và ngày tháng năm dạng đầy đủ, ví dụ:  14:15' PM, Thứ Sáu, 14/04/2006
    ''' </summary>
    ''' <param name="dateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getLongTimeDate(ByVal dateTime As Date) As String
        Try
            Dim day As String
            Select Case (Convert.ToString(dateTime.DayOfWeek))
                Case "Sunday"
                    day = "Chủ Nhật"
                Case "Monday"
                    day = "Thứ Hai"
                Case "Tuesday"
                    day = "Thứ Ba"
                Case "Wednesday"
                    day = "Thứ Tư"
                Case "Thursday"
                    day = "Thứ Năm"
                Case "Friday"
                    day = "Thứ Sáu"
                Case "Saturday"
                    day = "Thứ Bảy"
                Case Else
                    day = ""
            End Select
            day = dateTime.ToString("HH:mm") & "' " & ", " & day & ", " & dateTime.ToString("dd/MM/yyyy")  'Khong co PM/AM
            Return day
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function addMunite(ByVal startTime As String, ByVal munite As Integer) As String
        Dim v_date As Date
        Dim v_hour As Integer
        Dim v_munite As Integer
        Dim v_second As Integer
        Dim element() As String = startTime.Split(":")
        v_hour = Convert.ToInt32(element(0))
        v_munite = Convert.ToInt32(element(1))
        v_second = Convert.ToInt32(element(2))

        v_date = New Date(2008, 8, 10, v_hour, v_munite, v_second)
        v_date = v_date.AddMinutes(munite)
        Return addZero(v_date.Hour) + ":" + addZero(v_date.Minute) + ":" + addZero(v_date.Second)

    End Function

    Public Shared Function addZero(ByVal param As Integer) As String
        If param >= 0 And param <= 9 Then
            Return "0" & param
        Else : Return param.ToString
        End If
    End Function

    ''' <summary>
    ''' Hàm này trả về ngày cuối cùng trong tháng
    ''' </summary>
    ''' <param name="month"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLastDate(ByVal month As Integer, ByVal x As Boolean) As Integer
        Try
            Dim day As Integer = 0
            Select Case (month)
                Case 1, 3, 5, 7, 8, 12
                    day = 31
                Case 2
                    If x Then
                        day = 29
                    Else : day = 28
                    End If
                Case 4, 6, 9, 11
                    day = 30
                Case Else
                    day = 0
            End Select
            Return day
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Shared Function GetQuarter(ByVal fromDate As DateTime) As Integer
        Return ((fromDate.Month - 1) \ 3) + 1
    End Function

    Public Shared Function StringToDateTime(ByVal sDate As String, ByVal sFormat As String) As DateTime
        Try
            Dim oDate As DateTime = DateTime.ParseExact(sDate, sFormat, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None)
            Return oDate
        Catch ex As Exception
            Return MinSqlDateTime()
        End Try
    End Function

    Public Shared Function DateTimeToString(ByVal sDate As DateTime, ByVal sFormat As String) As String
        Try
            Dim oDate As String = sDate.ToString(sFormat, System.Globalization.CultureInfo.InvariantCulture)
            Return oDate
        Catch ex As Exception
            Return ""
        End Try
    End Function




#Region "---> Months <---"
    ''' <summary>
    ''' Hàm trả về ngày đầu tiên của tháng - Từ năm và tháng truyền vào
    ''' </summary>
    ''' <param name="pMonth">Tháng truyền vào - Kiểu nguyên</param>
    ''' <param name="pYear">Năm truyền vào - Kiểu nguyên</param>
    ''' <returns>Ngày đầu tiên của tháng</returns>
    ''' <remarks></remarks>
    Public Shared Function GetStartOfMonth(ByVal pMonth As Integer, ByVal pYear As Integer) As DateTime
        Return New DateTime(pYear, pMonth, 1, 0, 0, 0, 0)
    End Function

    ''' <summary>
    ''' Hàm trả về ngày cuối cùng của tháng - Từ năm và tháng truyền vào
    ''' </summary>
    ''' <param name="pMonth">Tháng truyền vào - Kiểu nguyên</param>
    ''' <param name="pYear">Năm truyền vào - Kiểu nguyên</param>
    ''' <returns>Ngày cuối cùng của tháng</returns>
    ''' <remarks></remarks>
    Public Shared Function GetEndOfMonth(ByVal pMonth As Integer, ByVal pYear As Integer) As DateTime
        Return New DateTime(pYear, pMonth, DateTime.DaysInMonth(pYear, pMonth), 23, 59, 59, 999)
    End Function
    Public Shared Function GetDateEndOfMonth(ByVal pMonth As Integer, ByVal pYear As Integer) As DateTime
        Return New DateTime(pYear, pMonth, DateTime.DaysInMonth(pYear, pMonth), 0, 0, 0, 0)
    End Function
    ''' <summary>
    ''' Hàm trả về ngày đầu tiên của tháng trước tháng hiện thời - Từ ngày hiện thời hệ thống (NOW)
    ''' </summary>
    ''' <returns>Ngày đầu tiên của tháng trước tháng hiện thời</returns>
    ''' <remarks></remarks>
    Public Shared Function GetStartOfLastMonth() As DateTime
        If DateTime.Now.Month = 1 Then
            Return GetStartOfMonth(12, DateTime.Now.Year - 1)
        Else
            Return GetStartOfMonth(DateTime.Now.Month - 1, DateTime.Now.Year)
        End If
    End Function

    ''' <summary>
    ''' Hàm trả về ngày đầu tiên của tháng trước tháng hiện thời - Từ ngày hiện thời truyền vào
    ''' </summary>
    ''' <param name="pDateTimeNow">Ngày hiện thời truyền vào</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetStartOfLastMonth(ByVal pDateTimeNow As DateTime) As DateTime
        If pDateTimeNow.Month = 1 Then
            Return GetStartOfMonth(12, pDateTimeNow.Year - 1)
        Else
            Return GetStartOfMonth(pDateTimeNow.Month - 1, pDateTimeNow.Year)
        End If
    End Function

    ''' <summary>
    ''' Hàm trả về ngày cuối cùng của tháng trước tháng hiện thời - Từ ngày hiện thời hệ thống (NOW)
    ''' </summary>
    ''' <returns>Ngày cuối cùng của tháng trước tháng hiện thời</returns>
    ''' <remarks></remarks>
    Public Shared Function GetEndOfLastMonth() As DateTime
        If DateTime.Now.Month = 1 Then
            Return GetEndOfMonth(12, DateTime.Now.Year - 1)
        Else
            Return GetEndOfMonth(DateTime.Now.Month - 1, DateTime.Now.Year)
        End If
    End Function

    ''' <summary>
    ''' Hàm trả về ngày cuối cùng của tháng trước tháng hiện thời - Từ ngày hiện thời truyền vào
    ''' </summary>
    ''' <param name="pDateTimeNow">Ngày hiện thời truyền vào</param>
    ''' <returns>Ngày cuối cùng của tháng trước tháng theo ngày truyền vào</returns>
    ''' <remarks></remarks>
    Public Shared Function GetEndOfLastMonth(ByVal pDateTimeNow As DateTime) As DateTime
        If pDateTimeNow.Month = 1 Then
            Return GetEndOfMonth(12, pDateTimeNow.Year - 1)
        Else
            Return GetEndOfMonth(pDateTimeNow.Month - 1, pDateTimeNow.Year)
        End If
    End Function

    ''' <summary>
    ''' Hàm trả về ngày đầu tiên của của tháng hiện thời
    ''' </summary>
    ''' <returns>Ngày đầu tiên của của tháng hiện thời</returns>
    ''' <remarks></remarks>
    Public Shared Function GetStartOfCurrentMonth() As DateTime
        Return GetStartOfMonth(DateTime.Now.Month, DateTime.Now.Year)
    End Function

    ''' <summary>
    ''' Hàm trả về ngày đầu tiên của của tháng hiện thời theo ngày truyền vào
    ''' </summary>
    ''' <param name="pDateTimeNow">Ngày hiện thời truyền vào</param>
    ''' <returns>Ngày đầu tiên của tháng</returns>
    ''' <remarks></remarks>
    Public Shared Function GetStartOfCurrentMonth(ByVal pDateTimeNow As DateTime) As DateTime
        Return GetStartOfMonth(pDateTimeNow.Month, pDateTimeNow.Year)
    End Function

    ''' <summary>
    ''' Hàm trả về ngày cuối cùng của của tháng hiện thời
    ''' </summary>
    ''' <returns>Ngày cuối cùng của của tháng hiện thời</returns>
    ''' <remarks></remarks>
    Public Shared Function GetEndOfCurrentMonth() As DateTime
        Return GetEndOfMonth(DateTime.Now.Month, DateTime.Now.Year)
    End Function

    ''' <summary>
    ''' Hàm trả về ngày cuối cùng của của tháng hiện thời
    ''' </summary>
    ''' <param name="pDateTimeNow">Ngày truyền vào</param>
    ''' <returns>Ngày cuối cùng của của tháng hiện thời theo ngày truyền vào</returns>
    ''' <remarks></remarks>
    Public Shared Function GetEndOfCurrentMonth(ByVal pDateTimeNow As DateTime) As DateTime
        Return GetEndOfMonth(pDateTimeNow.Month, pDateTimeNow.Year)
    End Function

    ''' <summary>
    ''' Hàm trả về ngày đầu tiên của Tháng - Từ ngày truyền vào
    ''' </summary>
    ''' <param name="pDateTimeNow">Ngày truyền vào</param>
    ''' <returns>Ngày đầu tiên của tháng</returns>
    ''' <remarks></remarks>
    Public Shared Function GetBeginDateOfMonthOfDate(ByVal pDateTimeNow As DateTime) As DateTime
        Return New DateTime(pDateTimeNow.Year, pDateTimeNow.Month, 1, 0, 0, 0, 0)
    End Function

    ''' <summary>
    ''' Ngày cuối cùng của tháng từ ngày truyền vào
    ''' </summary>
    ''' <param name="pDateTimeNow">Ngày truyền vào</param>
    ''' <returns>Ngày cuối cùng của tháng</returns>
    ''' <remarks></remarks>
    Public Shared Function GetEndDateOfMonthOfDate(ByVal pDateTimeNow As DateTime) As DateTime
        Dim firstDayOfTheMonth As DateTime = New DateTime(pDateTimeNow.Year, pDateTimeNow.Month, 1)
        Return firstDayOfTheMonth.AddMonths(1).AddDays(-1)
    End Function


#End Region

End Class
