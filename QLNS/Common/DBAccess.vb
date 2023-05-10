Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO

''' <summary>
''' Class chứa các hàm làm việc với SQLClient
''' Author: Nguyễn Thị Thuỳ Giang
''' Date: 3/8/2008
''' </summary>
''' <remarks></remarks>
Public Class DBAccess

    Private connectionString As String

    Public Sub New()
        connectionString = QLNS_CONNSTR
    End Sub

    Public Sub New(ByVal connString As String)
        connectionString = connString
    End Sub

#Region "Management Methods"

    Public Function getConnection() As SqlConnection
        Try
            Return New SqlConnection(connectionString)
        Catch sqlEx As SqlException
            Throw (New Exception("Không kết nối được với CSDL:" & sqlEx.Message))
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện kiểm tra xem chuỗi kết nối có truy xuất được CSDL không.
    ''' </summary>
    ''' <returns>True: Connect success. False: Failed to connect to data source</returns>
    ''' <remarks></remarks>
    Public Function CheckConnection() As Boolean
        Dim _connection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(connectionString)
        Try
            If (_connection.State <> System.Data.ConnectionState.Open) Then _connection.Open()
            Return True
        Catch ex As Exception
            Return False
        Finally
            If (_connection.State <> System.Data.ConnectionState.Closed) Then _connection.Close()
        End Try
    End Function

    'Tuong duong voi ham = getConnection
    ''' <summary>
    ''' Function is open connections - Mở kết nối cơ sở dữ liệu
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSqlConnection() As SqlConnection
        Dim _connection As SqlConnection = New SqlConnection(connectionString)
        If (_connection.State <> ConnectionState.Open) Then
            Try
                _connection.Open()
            Catch ex As Exception
                MessageBox.Show("Lost Connection!" + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return Nothing
            End Try
        End If
        Return _connection
    End Function


    Public Function openConnection(ByVal connString As String) As SqlConnection
        Try
            Dim mySqlconnection As SqlConnection = New SqlConnection(connString)
            mySqlconnection.Open()
            Return mySqlconnection
        Catch sqlEx As SqlException
            Throw New Exception(sqlEx.Message)
        End Try
    End Function
      
    Public Function openConnection() As SqlConnection
        Try
            Dim mySqlconnection As SqlConnection = New SqlConnection(connectionString)
            mySqlconnection.Open()
            Return mySqlconnection
        Catch sqlEx As SqlException
            Throw New Exception(sqlEx.Message)
        End Try
    End Function
     

    Public Sub closeConnection(ByVal conn As SqlConnection)
        Try
            If (Not (conn Is DBNull.Value) Or conn.State = ConnectionState.Open) Then
                conn.Close()
                conn.Dispose()
            End If
        Catch sqlEx As SqlException
            Throw New Exception(sqlEx.Message)
        End Try
    End Sub

    Public Function getDataSet(ByVal strSQL As String) As DataSet
        Dim my_Base As DBAccess = New DBAccess
        Try
            Dim conn As SqlConnection = my_Base.openConnection(connectionString)
            Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, conn)
            Dim ds As DataSet = New DataSet
            da.Fill(ds)
            da.Dispose()
            my_Base.closeConnection(conn)
            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message & strSQL)
        End Try
    End Function

    Public Function getReader(ByVal cmd As SqlCommand) As SqlDataReader
        Dim my_Base As New DBAccess
        Try
            Dim myconn As SqlConnection = New SqlConnection(connectionString)
            cmd.Connection = myconn
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Return reader
        Catch ex As Exception
            Throw New Exception(ex.Message & cmd.CommandText)
        End Try
    End Function
        
    Public Function getReader(ByVal cmd As SqlCommand, ByVal conn As SqlConnection) As SqlDataReader
        Try
            cmd.Connection = conn
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Return reader
        Catch ex As Exception
            Throw New Exception(ex.Message & cmd.CommandText)
        End Try
    End Function

    Public Function getDataTable(ByVal strSQL As String) As DataTable
        Return getDataSet(strSQL).Tables(0)
    End Function

    Public Function getDataSet(ByVal strSQL As String, ByVal Parameters As SqlParameter()) As DataSet
        Dim my_Base As DBAccess = New DBAccess
        Try
            Dim myconn As SqlConnection = my_Base.openConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSQL, myconn)
            Dim par As SqlParameter
            cmd.Parameters.Clear()
            For Each par In Parameters
                cmd.Parameters.Add(par)
            Next
            Dim myda As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myds As DataSet = New DataSet
            myda.Fill(myds)
            my_Base.closeConnection(myconn)
            Return myds

        Catch ex As Exception
            Throw New Exception(ex.Message & strSQL)
        End Try
    End Function

    Public Function getDataTable(ByVal strSQL As String, ByVal Parameters() As SqlParameter) As DataTable
        Return getDataSet(strSQL, Parameters).Tables(0)
    End Function
 
    Public Function getDataSet(ByVal cmd As SqlCommand) As DataSet
        Dim my_Base As DBAccess = New DBAccess
        Try
            Dim conn As SqlConnection = my_Base.openConnection(connectionString)
            cmd.Connection = conn
            Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim ds As DataSet = New DataSet
            da.Fill(ds)
            my_Base.closeConnection(conn)
            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message & cmd.CommandText)
        End Try
    End Function

    Public Function getDataTable(ByVal cmd As SqlCommand) As DataTable
        Return getDataSet(cmd).Tables(0)
    End Function

    Public Sub executeSQL(ByVal strSQL As String)
        Dim my_Base As DBAccess = New DBAccess
        Dim conn As SqlConnection = my_Base.openConnection(connectionString)
        Dim tran As SqlTransaction = conn.BeginTransaction
        Try
            Dim cmd As SqlCommand = New SqlCommand(strSQL, conn)
            cmd.Transaction = tran
            cmd.ExecuteNonQuery()
            tran.Commit()
            cmd.Dispose()
        Catch ex As Exception
            tran.Rollback()
            Throw New Exception(ex.Message & strSQL)
        Finally
            my_Base.closeConnection(conn)
        End Try
    End Sub

    Public Sub executeSQL(ByVal strSQL As String, ByVal parameter() As SqlParameter)
        Dim my_Base As DBAccess = New DBAccess
        Try
            Dim conn As SqlConnection = my_Base.openConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSQL, conn)
            cmd.Parameters.Clear()
            Dim par As SqlParameter
            For Each par In parameter
                cmd.Parameters.Add(par)
            Next
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            my_Base.closeConnection(conn)
        Catch ex As Exception
            Throw New Exception(ex.Message & strSQL)
        End Try
    End Sub

    Public Sub executeSQL(ByVal strSQL As String, ByVal paraList As IList)
        Dim my_Base As DBAccess = New DBAccess
        Try
            Dim conn As SqlConnection = my_Base.openConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSQL, conn)
            Dim par As SqlParameter
            For Each par In paraList
                cmd.Parameters.Add(par)
            Next
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            my_Base.closeConnection(conn)
        Catch ex As Exception
            Throw New Exception(ex.Message & strSQL)
        End Try
    End Sub

    Public Sub executeSQL(ByVal cmd As SqlCommand)
        Dim my_Base As DBAccess = New DBAccess
        Try
            Dim conn As SqlConnection = my_Base.openConnection(connectionString)
            cmd.Connection = conn
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            my_Base.closeConnection(conn)
        Catch ex As Exception
            Throw New Exception(ex.Message & cmd.CommandText)
        End Try
    End Sub
 
    Public Sub executeSQL(ByVal strSQL As String, ByVal conn As SqlConnection)
        Try
            Dim cmd As SqlCommand = New SqlCommand(strSQL, conn)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
        Catch ex As Exception
            Throw New Exception(ex.Message & strSQL)
        End Try
    End Sub

    Public Function SelectDBRows(ByVal strSql As String) As DataTable
        Try
            Dim conn As SqlConnection = New SqlConnection(connectionString)
            Dim dt As DataTable = New DataTable
            Dim da As SqlDataAdapter = New SqlDataAdapter
            conn.Open()
            da.SelectCommand = New SqlCommand(strSql, conn)
            da.Fill(dt)
            conn.Close()
            da.Dispose()
            conn.Dispose()
            Return dt
        Catch ex As Exception
            My_MessageBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function getNumber(ByVal strSql As String) As Integer
        Try
            Dim num As Object
            Dim conn As SqlConnection = New SqlConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSql, conn)
            conn.Open()
            num = CType(cmd.ExecuteScalar, Integer)
            conn.Close()
            conn.Dispose()
            Return num
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function getDouble(ByVal strSql As String) As Double
        Try
            Dim num As Object
            Dim conn As SqlConnection = New SqlConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSql, conn)
            conn.Open()
            num = CType(cmd.ExecuteScalar, Double)
            conn.Close()
            conn.Dispose()
            Return num
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function getBigNumber(ByVal strSql As String) As Long
        Try
            Dim num As Object
            Dim conn As SqlConnection = New SqlConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSql, conn)
            conn.Open()
            num = CType(cmd.ExecuteScalar, Long)
            conn.Close()
            conn.Dispose()
            Return num
        Catch ex As Exception
            Return 0
        End Try
    End Function
  
    Public Function getString(ByVal strSql As String) As String
        Try
            Dim str As String
            Dim conn As SqlConnection = New SqlConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSql, conn)
            conn.Open()
            str = CType(cmd.ExecuteScalar, String)
            conn.Close()
            conn.Dispose()
            Return str
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getDateTime(ByVal strSql As String) As DateTime
        Try
            Dim str As String
            Dim conn As SqlConnection = New SqlConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(strSql, conn)
            conn.Open()
            str = CType(cmd.ExecuteScalar, DateTime)
            conn.Close()
            conn.Dispose()
            Return str
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getNumber(ByVal SelectQuery As String, ByVal StringConn As String) As Long
        Dim conn As SqlConnection
        Try
            Dim valueResult As Object
            conn = New SqlConnection(StringConn)
            Dim getComm As New SqlCommand(SelectQuery, conn)
            conn.Open()
            valueResult = getComm.ExecuteScalar()
            conn.Close()
            getComm.Dispose()
            Return CLng(valueResult)
        Catch ex As Exception
            Return 0
        End Try
        Return 0
    End Function

    Public Function getString(ByVal SelectQuery As String, ByVal StringConn As String) As String
        Dim Conn As SqlConnection
        Try
            Conn = New SqlConnection(StringConn)
            Dim valueResult As Object
            Conn.Open()
            Dim getComm As New SqlCommand(SelectQuery, Conn)
            valueResult = getComm.ExecuteScalar()
            Conn.Close()
            getComm.Dispose()
            Return CStr(valueResult)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

End Class
