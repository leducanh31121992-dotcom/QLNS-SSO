Public Class DbCommon
    ''' <summary>
    ''' Open connection - Hàm thực hiện mở kết nối cơ sở dữ liệu
    ''' </summary>
    ''' <returns>Connection</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSqlConnection() As System.Data.SqlClient.SqlConnection
        Dim strConn As String = QLNS_CONNSTR
        Dim _connection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(strConn)
        If (_connection.State <> System.Data.ConnectionState.Open) Then
            Try
                _connection.Open()
            Catch ex As Exception
                MessageBox.Show("Lỗi kết nối cơ sở dữ liệu: " + ex.Message.ToString(), "Connection...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return Nothing
            End Try
        End If
        Return _connection
    End Function

    ''' <summary>
    ''' Open connection - Hàm thực hiện mở kết nối cơ sở dữ liệu
    ''' </summary>
    ''' <param name="strConn">Chuỗi connection truyền vào</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSqlConnection(ByVal strConn As String) As System.Data.SqlClient.SqlConnection
        Dim _connection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(strConn)
        If (_connection.State <> System.Data.ConnectionState.Open) Then
            Try
                _connection.Open()
            Catch ex As Exception
                MessageBox.Show("Lỗi kết nối cơ sở dữ liệu: " + ex.Message.ToString(), "Connection...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return Nothing
            End Try
        End If
        Return _connection
    End Function

    ''' <summary>
    ''' Close connection - Hàm đóng kết nối cơ sở dữ liệu
    ''' </summary>
    ''' <param name="_connection"></param>
    ''' <remarks></remarks>
    Public Shared Sub CloseConnection(ByVal _connection As System.Data.SqlClient.SqlConnection)
        If (_connection IsNot Nothing) Then
            If (_connection.State <> System.Data.ConnectionState.Closed) Then
                Try
                    _connection.Close()
                Catch ex As Exception
                    MessageBox.Show("Lỗi đóng kết nối cơ sở dữ liệu: " + ex.Message.ToString(), "Closed connection...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
            End If
        End If
    End Sub
End Class
