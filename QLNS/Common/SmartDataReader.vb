Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class SmartDataReader

    Private defaultDate As Date
    Private reader As SqlDataReader

    Public Sub New()

    End Sub

    Public Sub New(ByVal reader As SqlDataReader)
        defaultDate = Date.MinValue
        Me.reader = reader
    End Sub

    Public Function getInt32(ByVal column As String) As Integer
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CInt(0), CInt(reader(column)))
        Catch ex As Exception
            Return CInt(0)
        End Try
    End Function

    Public Function getInt64(ByVal column As String) As Long
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CType(0, Long), CType(reader(column), Long))
        Catch ex As Exception
            Return CType(0, Long)
        End Try
    End Function

    Public Function getInt16(ByVal column As String) As Short
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CShort(0), CShort(reader(column)))
        Catch ex As Exception
            Return CShort(0)
        End Try
    End Function

    Public Function getByte(ByVal column As String) As Byte
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CByte(0), CByte(reader(column)))
        Catch ex As Exception
            Return CByte(0)
        End Try
    End Function

    Public Function getFloat(ByVal column As String) As Double
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CDbl(0), CDbl(reader(column)))
        Catch ex As Exception
            Return CDbl(0)
        End Try
    End Function

    Public Function getBoolean(ByVal column As String) As Boolean
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CBool(0), CBool(reader(column)))
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function getString(ByVal column As String) As String
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), "", reader(column).ToString)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getDatetime(ByVal column As String) As Date
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CDate(defaultDate), CDate(reader(column)))
        Catch ex As Exception
            Return defaultDate
        End Try
    End Function

    Public Function getDecimal(ByVal column As String) As Decimal
        Try
            Return IIf(reader.IsDBNull(reader.GetOrdinal(column)), CDec(0), CDec(reader(column)))
        Catch ex As Exception
            Return CDec(0)
        End Try
    End Function

    Public Function Read() As Boolean
        Return reader.Read
    End Function

    Public Sub disposeReader(ByVal reader As SqlDataReader)
        If (Not reader Is DBNull.Value Or Not reader.IsClosed) Then
            reader.Close()
        End If

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
