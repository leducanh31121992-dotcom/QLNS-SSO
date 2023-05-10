''' <summary>
''' Write by: Dương Văn Chữ - Tháng 11 năm 2008
''' The SqlHelper class is intended to encapsulate high performance, scalable best practices for common uses of SqlClient
''' </summary>
''' <remarks></remarks>
Public Class SoftSqlHelper

#Region "-----> AttachParameters <-----"
    ''' <summary>
    ''' This method is used to attach array of SqlParameters to a SqlCommand.
    ''' This method will assign a value of DbNull to any parameter with a direction of InputOutput and a value of null. 
    ''' This behavior will prevent default values from being used, but
    ''' this will be the less common case than an intended pure output parameter (derived as InputOutput)
    ''' where the user provided no input value.
    ''' </summary>
    ''' <param name="_command">The command to which the parameters will be added</param>
    ''' <param name="_commandParameters">An array of SqlParameters to be added to command</param>
    ''' <remarks></remarks>
    Private Shared Sub AttachParameters(ByVal _command As System.Data.SqlClient.SqlCommand, ByVal _commandParameters As System.Data.SqlClient.SqlParameter())
        If (_command Is Nothing) Then Throw New ArgumentNullException("command")
        If (_commandParameters IsNot Nothing) Then
            For Each _param As System.Data.SqlClient.SqlParameter In _commandParameters
                If (_param IsNot Nothing) Then
                    'Check for derived output value with no value assigned
                    If ((_param.Direction = ParameterDirection.InputOutput Or _param.Direction = ParameterDirection.Input) And (_param.Value Is Nothing)) Then
                        _param.Value = DBNull.Value
                    End If
                    _command.Parameters.Add(_param)
                End If
            Next
        End If
    End Sub

    Private Shared Sub AttachParameters(ByVal _command As System.Data.SqlClient.SqlCommand, ByVal _commandParameters As System.Data.SqlClient.SqlParameterCollection)
        If (_command Is Nothing) Then Throw New ArgumentNullException("command")
        If (_commandParameters IsNot Nothing) Then
            For Each _param As System.Data.SqlClient.SqlParameter In _commandParameters
                If ((_param.Direction = ParameterDirection.InputOutput) And (_param.Value Is Nothing)) Then
                    _param.Value = DBNull.Value
                End If
                _command.Parameters.Add(_param)
            Next
        End If
    End Sub
#End Region

#Region "-----> AssignParameterValues <-----"
    ''' <summary>
    ''' This method assigns dataRow column values to an array of SqlParameters
    ''' </summary>
    ''' <param name="_commandParameters">Array of SqlParameters to be assigned values</param>
    ''' <param name="_dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <remarks></remarks>
    Private Shared Sub AssignParameterValues(ByVal _commandParameters As System.Data.SqlClient.SqlParameter(), ByVal _dataRow As DataRow)
        If ((_commandParameters Is Nothing) Or (_dataRow Is Nothing)) Then
            'Do nothing if we get no data
            Return
        End If
        Dim i As Integer = 0
        ' Set the parameters values
        For Each _param As System.Data.SqlClient.SqlParameter In _commandParameters
            ' Check the parameter name
            If ((_param.ParameterName Is Nothing) Or (_param.ParameterName.Length <= 1)) Then
                Throw New Exception(String.Format("Please provide a valid parameter name on the parameter #{0}, the ParameterName property has the following value: '{1}'.", i, _param.ParameterName))
            End If
            If (_dataRow.Table.Columns.IndexOf(_param.ParameterName.Substring(1)) <> -1) Then
                _param.Value = _dataRow(_param.ParameterName.Substring(1))
            End If
            i = i + 1
        Next
    End Sub

    ''' <summary>
    ''' This method assigns an array of values to an array of SqlParameters
    ''' </summary>
    ''' <param name="_commandParameters">Array of SqlParameters to be assigned values</param>
    ''' <param name="_parameterValues">Array of objects holding the values to be assigned</param>
    ''' <remarks></remarks>
    Private Shared Sub AssignParameterValues(ByVal _commandParameters As System.Data.SqlClient.SqlParameter(), ByVal _parameterValues As Object())
        If ((_commandParameters Is Nothing) Or (_parameterValues Is Nothing)) Then
            'Do nothing if we get no data
            Return
        End If

        'We must have the same number of values as we pave parameters to put them in
        If (_commandParameters.Length <> _parameterValues.Length) Then
            Throw New ArgumentException("Parameter count does not match Parameter Value count.")
        End If

        ' Iterate through the SqlParameters, assigning the values from the corresponding position in the value array
        Dim j As Integer = _commandParameters.Length
        For i As Integer = 0 To j - 1
            'If the current array value derives from IDbDataParameter, then assign its Value property
            If (TypeOf _parameterValues(i) Is System.Data.IDbDataParameter) Then
                Dim paramInstance As System.Data.IDbDataParameter = CType(_parameterValues(i), IDbDataParameter)
                If (paramInstance.Value Is Nothing) Then
                    _commandParameters(i).Value = DBNull.Value
                Else
                    _commandParameters(i).Value = paramInstance.Value
                End If
            ElseIf (_parameterValues(i) Is Nothing) Then
                _commandParameters(i).Value = DBNull.Value
            Else
                _commandParameters(i).Value = _parameterValues(i)
            End If
        Next
    End Sub



    'for (int i = 0, j = commandParameters.Length; i < j; i++)
    '            {
    '                // If the current array value derives from IDbDataParameter, then assign its Value property
    '				if (parameterValues[i] is IDbDataParameter)
    '				{
    '					IDbDataParameter paramInstance = (IDbDataParameter)parameterValues[i];
    '					if( paramInstance.Value == null )
    '					{
    '						commandParameters[i].Value = DBNull.Value; 
    '					}
    '					else
    '					{
    '						commandParameters[i].Value = paramInstance.Value;
    '					}
    '				}
    '				else if (parameterValues[i] == null)
    '				{
    '					commandParameters[i].Value = DBNull.Value;
    '				}
    '				else
    '				{
    '					commandParameters[i].Value = parameterValues[i];
    '				}
    '            }
    '        }




#End Region

#Region "-----> PrepareCommand <-----"
    ''' <summary>
    ''' This method opens (if necessary) and assigns a connection, transaction, command type and parameters to the provided command
    ''' </summary>
    ''' <param name="_command">The SqlCommand to be prepared</param>
    ''' <param name="_connection">A valid SqlConnection, on which to execute this command</param>
    ''' <param name="_transaction">A valid SqlTransaction, or 'nothing'</param>
    ''' <param name="_commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="_commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="_commandParameters">An array of SqlParameters to be associated with the command or 'nothing' if no parameters are required</param>
    ''' <param name="mustCloseConnection"><c>true</c> if the connection was opened by the method, otherwose is false</param>
    ''' <remarks></remarks>
    Private Shared Sub PrepareCommand(ByVal _command As System.Data.SqlClient.SqlCommand, ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal _transaction As System.Data.SqlClient.SqlTransaction, ByVal _commandType As System.Data.CommandType, ByVal _commandText As String, ByVal _commandParameters As System.Data.SqlClient.SqlParameter(), ByVal mustCloseConnection As Boolean)
        If (_command Is Nothing) Then Throw New ArgumentNullException("_command")
        If (_commandText Is Nothing Or _commandText.Length = 0) Then Throw New ArgumentNullException("_commandText")
        'If the provided connection is not open, we will open it
        Try
            If (_connection.State <> ConnectionState.Open) Then
                mustCloseConnection = True
                _connection.Open()
            Else
                mustCloseConnection = False
            End If
        Catch ex As Exception
            Throw New ArgumentNullException("_connection")
        End Try
        'Associate the connection with the command
        _command.Connection = _connection
        'Set the command text (stored procedure name or SQL statement)
        _command.CommandText = _commandText

        'If we were provided a transaction, assign it
        If (_transaction IsNot Nothing) Then
            If (_transaction.Connection Is Nothing) Then
                Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
            End If
            _command.Transaction = _transaction
        End If
        'Set the command type
        _command.CommandType = _commandType
        'Attach the command parameters if they are provided
        If (_commandParameters IsNot Nothing) Then
            AttachParameters(_command, _commandParameters)
        End If
        Return
    End Sub

    ''' <summary>
    ''' This method opens (if necessary) and assigns a connection, transaction, command type and parameters to the provided command
    ''' </summary>
    ''' <param name="_command">The SqlCommand to be prepared</param>
    ''' <param name="_connection">A valid SqlConnection, on which to execute this command</param>
    ''' <param name="_transaction">A valid SqlTransaction, or 'nothing'</param>
    ''' <param name="_commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="_commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="_commandParameters">An array of SqlParameters to be associated with the command or 'nothing' if no parameters are required</param>
    ''' <remarks></remarks>
    Private Shared Sub PrepareCommand(ByVal _command As System.Data.SqlClient.SqlCommand, ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal _transaction As System.Data.SqlClient.SqlTransaction, ByVal _commandType As System.Data.CommandType, ByVal _commandText As String, ByVal _commandParameters As System.Data.SqlClient.SqlParameter())
        If (_command Is Nothing) Then Throw New ArgumentNullException("command")
        If (_commandText Is Nothing Or _commandText.Length = 0) Then Throw New ArgumentNullException("commandText")
        'If the provided connection is not open, we will open it
        If (_connection.State <> ConnectionState.Open) Then
            Try
                _connection.Open()
            Catch ex As Exception
                Throw New ArgumentNullException("connection")
            End Try
        End If
        'Associate the connection with the command
        _command.Connection = _connection
        'Set the command text (stored procedure name or SQL statement)
        _command.CommandText = _commandText
        'If we were provided a transaction, assign it
        If (_transaction IsNot Nothing) Then
            If (_transaction.Connection Is Nothing) Then
                Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
            End If
            _command.Transaction = _transaction
        End If
        'Set the command type
        _command.CommandType = _commandType
        'Attach the command parameters if they are provided
        If (_commandParameters IsNot Nothing) Then
            AttachParameters(_command, _commandParameters)
        End If
        Return
    End Sub

    Private Shared Sub PrepareCommand(ByVal _command As System.Data.SqlClient.SqlCommand, ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal _transaction As System.Data.SqlClient.SqlTransaction, ByVal _commandType As System.Data.CommandType, ByVal _commandText As String, ByVal _commandParameters As System.Data.SqlClient.SqlParameterCollection, ByVal mustCloseConnection As Boolean)
        If (_command Is Nothing) Then Throw New ArgumentNullException("_command")
        If (_commandText Is Nothing Or _commandText.Length = 0) Then Throw New ArgumentNullException("_commandText")
        'If the provided connection is not open, we will open it
        Try
            If (_connection.State <> ConnectionState.Open) Then
                mustCloseConnection = True
                _connection.Open()
            Else
                mustCloseConnection = False
            End If
        Catch ex As Exception
            Throw New ArgumentNullException("_connection")
        End Try
        'Associate the connection with the command
        _command.Connection = _connection
        'Set the command text (stored procedure name or SQL statement)
        _command.CommandText = _commandText

        'If we were provided a transaction, assign it
        If (_transaction IsNot Nothing) Then
            If (_transaction.Connection Is Nothing) Then
                Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
            End If
            _command.Transaction = _transaction
        End If
        _command.CommandType = _commandType

        If (_commandParameters IsNot Nothing) Then
            AttachParameters(_command, _commandParameters)
        End If
        Return
    End Sub

    Private Shared Sub PrepareCommand(ByVal _command As System.Data.SqlClient.SqlCommand, ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal _transaction As System.Data.SqlClient.SqlTransaction, ByVal _commandType As System.Data.CommandType, ByVal _commandText As String, ByVal _commandParameters As System.Data.SqlClient.SqlParameterCollection)
        If (_command Is Nothing) Then Throw New ArgumentNullException("command")
        If (_commandText Is Nothing Or _commandText.Length = 0) Then Throw New ArgumentNullException("commandText")
        'If the provided connection is not open, we will open it
        If (_connection.State <> ConnectionState.Open) Then
            Try
                _connection.Open()
            Catch ex As Exception
                Throw New ArgumentNullException("connection")
            End Try
        End If
        'Associate the connection with the command
        _command.Connection = _connection
        'Set the command text (stored procedure name or SQL statement)
        _command.CommandText = _commandText

        'If we were provided a transaction, assign it
        If (_transaction IsNot Nothing) Then
            If (_transaction.Connection Is Nothing) Then
                Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
            End If
            _command.Transaction = _transaction
        End If
        _command.CommandType = _commandType

        If (_commandParameters IsNot Nothing) Then
            AttachParameters(_command, _commandParameters)
        End If
        Return
    End Sub
#End Region

#Region "-----> CreateParameter <-----"
    Public Shared Function CreateParameter(ByVal _paramName As String, ByVal _paramValue As Object, ByVal _paramDirection As System.Data.ParameterDirection) As System.Data.SqlClient.SqlParameter
        Dim _param As System.Data.SqlClient.SqlParameter
        If (TypeOf _paramValue Is System.Data.SqlTypes.SqlDateTime) Then
            If (CType(_paramValue, System.Data.SqlTypes.SqlDateTime).IsNull) Then
                'para = new MySqlParameter(paramName, null)
                _param = New SqlClient.SqlParameter(_paramName, Nothing)
            Else
                'para = new MySqlParameter(paramName, (DateTime)((MySqlDateTime)paramValue))
                _param = New SqlClient.SqlParameter(_paramName, CType(CType(_paramValue, System.Data.SqlTypes.SqlDateTime), DateTime))
            End If
        Else
            'para = new MySqlParameter(paramName, paramValue)
            _param = New SqlClient.SqlParameter(_paramName, _paramValue)
        End If
        _param.Direction = _paramDirection
        Return _param
    End Function
#End Region

#Region "-----> CreateCommand <-----"
    ''' <summary>
    ''' Simplify the creation of a Sql command object by allowing a stored procedure and optional parameters to be provided
    ''' </summary>
    ''' <param name="_connection">A valid SqlConnection object</param>
    ''' <param name="_spName">The name of the stored procedure</param>
    ''' <param name="_sourceColumns">An array of string to be assigned as the source columns of the stored procedure parameters</param>
    ''' <returns>A valid SqlCommand object</returns>
    ''' <remarks>e.g.: command as SqlCommand = CreateCommand(conn, "AddCustomer", "CustomerID", "CustomerName")</remarks>
    Public Shared Function CreateCommand(ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal _spName As String, ByVal _sourceColumns As String()) As System.Data.SqlClient.SqlCommand
        If (_connection Is Nothing) Then Throw New ArgumentNullException("_connection")
        If ((_spName Is Nothing) Or (_spName.Length = 0)) Then Throw New ArgumentNullException("_spName")
        'Create a SqlCommand
        Dim _command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(_spName, _connection)
        _command.CommandType = CommandType.StoredProcedure

        'If we receive parameter values, we need to figure out where they go
        If ((_sourceColumns IsNot Nothing) And (_sourceColumns.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(_connection, _spName)

            'Assign the provided source columns to these parameters based on parameter order

            For i As Integer = 0 To _sourceColumns.Length - 1
                commandParameters(i).SourceColumn = _sourceColumns(i)
            Next
            'Attach the discovered parameters to the SqlCommand object
            AttachParameters(_command, commandParameters)
        End If
        Return _command
    End Function
#End Region

#Region "-----> ExecuteDataset <-----"

    Public Shared Function ExecuteDataset(ByVal connection As System.Data.SqlClient.SqlConnection, ByVal commandType As System.Data.CommandType, ByVal commandText As String, ByVal commandParameters As System.Data.SqlClient.SqlParameterCollection) As DataSet
        If (connection Is Nothing) Then Return Nothing
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters, mustCloseConnection)
        Using mydap As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            'Fill the DataSet using default values for DataTable names, etc
            Try
                mydap.Fill(ds)
            Catch ex As Exception
                Throw ex
            End Try
            'Detach the SqlParameters from the command object, so they can be used again
            command.Parameters.Clear()
            If (mustCloseConnection) Then
                connection.Close()
            End If
            'Return the dataset
            Return ds
        End Using
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the database specified in the connection string
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>E.g:  Dim ds as DataSet = ExecuteDataset(connString, CommandType.StoredProcedure, "GetOrders")</remarks>
    Public Shared Function ExecuteDataset(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String) As DataSet
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteDataset(connectionString, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the database specified in the connection string using the provided parameters.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>E.g: Dim ds as DataSet = ExecuteDataset(connString, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))</remarks>
    Public Shared Function ExecuteDataset(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As System.Data.SqlClient.SqlParameter()) As DataSet
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        'Create & open a SqlConnection, and dispose of it after we are done
        Using connection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(connectionString)
            connection.Open()
            'Call the overload that takes a connection in place of the connection string
            Return ExecuteDataset(connection, commandType, commandText, commandParameters)
        End Using
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in 
    ''' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>This method provides no access to output parameters or the stored procedure's return value parameter
    ''' E.g:  Dim ds as DataSet = ExecuteDataset(connString, "GetOrders", 24, 36)</remarks>
    Public Shared Function ExecuteDataset(ByVal connectionString As String, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As DataSet
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")

        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As System.Data.SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)
            'Call the overload that takes an array of SqlParameters
            Return ExecuteDataset(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteDataset(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlConnection
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>Eg: Dim ds as DataSet = ExecuteDataset(conn, CommandType.StoredProcedure, "GetOrders")</remarks>
    Public Shared Function ExecuteDataset(ByVal _connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String) As DataSet
        'Pass through the call providing null for the set of SqlParameters
        Dim _commandParameters As SqlClient.SqlParameter() = Nothing
        Return ExecuteDataset(_connection, commandType, commandText, _commandParameters)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection using the provided parameters.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>Eg: Dim ds as DataSet = ExecuteDataset(conn, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))</remarks>
    Public Shared Function ExecuteDataset(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As DataSet
        If (connection Is Nothing) Then Return Nothing
        'Create a command and prepare it for execution
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters, mustCloseConnection)
        'Create the DataAdapter & DataSet
        Using mydap As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            'Fill the DataSet using default values for DataTable names, etc
            Try
                mydap.Fill(ds)
            Catch ex As Exception
                Throw ex
            End Try

            'Detach the SqlParameters from the command object, so they can be used again
            command.Parameters.Clear()
            If (mustCloseConnection) Then
                connection.Close()
            End If
            ' Return the dataset
            Return ds
        End Using
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ''' using the provided parameter values. This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Eg: Dim ds as DataSet = ExecuteDataset(conn, "GetOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteDataset(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As DataSet
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            ' Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)
            'Call the overload that takes an array of SqlParameters
            Return ExecuteDataset(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteDataset(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlTransaction
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim ds as DataSet = ExecuteDataset(trans, CommandType.StoredProcedure, "GetOrders")
    ''' </remarks>
    Public Shared Function ExecuteDataset(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String) As DataSet
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteDataset(transaction, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlTransaction using the provided parameters
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim ds As DataSet = ExecuteDataset(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteDataset(ByVal transaction As System.Data.SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As System.Data.SqlClient.SqlParameter()) As DataSet
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")

        'Create a command and prepare it for execution
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)

        'Create the DataAdapter & DataSet
        Using mydap As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            'Fill the DataSet using default values for DataTable names, etc
            mydap.Fill(ds)
            'Detach the SqlParameters from the command object, so they can be used again
            command.Parameters.Clear()
            'Return the dataset
            Return ds
        End Using
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified
    ''' SqlTransaction using the provided parameter values.  This method will query the database to discover the parameters for the
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Eg: Dim ds As DataSet = ExecuteDataset(trans, "GetOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteDataset(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As DataSet
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")

        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)
            'Call the overload that takes an array of SqlParameters
            Return ExecuteDataset(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteDataset(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function

#End Region

#Region "-----> ExecuteForSQL <-----"
    Public Shared Function ExecuteForSQL(ByVal connnection As System.Data.SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal commandParameters As System.Data.SqlClient.SqlParameterCollection) As DataTable
        If (connnection Is Nothing) Then Return Nothing
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand()
        PrepareCommand(command, connnection, Nothing, commandType, commandText, commandParameters)
        Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
            Dim mydb As DataTable
            Dim ds As DataSet = New DataSet()
            'Fill the DataSet using default values for DataTable names
            Try
                mydap.Fill(ds, "dsTableName")
                mydb = ds.Tables(0)
            Catch ex As Exception
                mydb = Nothing
            End Try
            'Detach the SqlParameters from the command object, so they can be used again
            command.Parameters.Clear()
            ds.Clear()
            'Return the datatable
            Return mydb
        End Using
    End Function

    Public Shared Function ExecuteForSQL(ByVal connnection As System.Data.SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As System.Data.SqlClient.SqlParameter()) As DataTable
        If (connnection Is Nothing) Then Return Nothing
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand()
        PrepareCommand(command, connnection, Nothing, commandType, commandText, commandParameters)
        Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
            Dim mydb As DataTable
            Dim ds As DataSet = New DataSet()
            'Fill the DataSet using default values for DataTable names
            Try
                mydap.Fill(ds, "dsTableName")
                mydb = ds.Tables(0)
            Catch ex As Exception
                mydb = Nothing
            End Try
            'Detach the SqlParameters from the command object, so they can be used again
            command.Parameters.Clear()
            ds.Clear()
            'Return the datatable
            Return mydb
        End Using
    End Function

    Public Shared Function ExecuteForSQL(ByVal strSQL As String) As DataTable
        Dim connnection As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
        Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(strSQL, connnection)
            Dim mydb As DataTable = New DataTable()
            mydap.Fill(mydb)
            If ((connnection IsNot Nothing) And (connnection.State <> System.Data.ConnectionState.Closed)) Then
                connnection.Close()
            End If
            Return mydb
        End Using
    End Function

    ''' <summary>
    ''' Hàm thực hiện câu lệnh SQL trả về DataTable
    ''' </summary>
    ''' <param name="strSQL">Câu lệnh Query</param>
    ''' <param name="strConn">Chuỗi kết nối csdl</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteForSQL(ByVal strSQL As String, ByVal strConn As String) As DataTable
        Dim connnection As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection(strConn)
        Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(strSQL, connnection)
            Dim mydb As DataTable = New DataTable()
            mydap.Fill(mydb)
            If ((connnection IsNot Nothing) And (connnection.State <> System.Data.ConnectionState.Closed)) Then
                connnection.Close()
            End If
            Return mydb
        End Using
    End Function
#End Region

#Region "-----> ExecuteNonQuery <-----"

    Public Shared Function ExecuteNonQuery(ByVal connection As System.Data.SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal commandParameters As System.Data.SqlClient.SqlParameterCollection) As Integer
        If (connection Is Nothing) Then Return 0
        'Create a command and prepare it for execution
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters, mustCloseConnection)
        'Finally, execute the command
        Dim retval As Integer = command.ExecuteNonQuery()
        command.Parameters.Clear()
        If (mustCloseConnection) Then
            connection.Close()
        End If
        Return retval
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns no resultset and takes no parameters) against the database specified in the connection string
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' Eg: Dim result as Integer = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders")
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String) As Integer
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteNonQuery(connectionString, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns no resultset) against the database specified in the connection string using the provided parameters
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' Eg: Dim result as Integer = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As System.Data.SqlClient.SqlParameter()) As Integer
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        'Create & open a SqlConnection, and dispose of it after we are done
        Using connection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(connectionString)
            connection.Open()
            'Call the overload that takes a connection in place of the connection string
            Return ExecuteNonQuery(connection, commandType, commandText, commandParameters)
        End Using
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns no resultset) against the database specified in 
    ''' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored prcedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Eg: Dim result As Integer = ExecuteNonQuery(connString, "PublishOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal connectionString As String, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As Integer
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As System.Data.SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'Call the overload that takes an array of SqlParameters
            Return ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns no resultset and takes no parameters) against the provided SqlConnection.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' Eg: Dim result as Integer = ExecuteNonQuery(conn, CommandType.StoredProcedure, "PublishOrders")
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal connection As System.Data.SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String) As Integer
        'Pass through the call providing null for the set of SqlParameters
        Dim commandParameters As System.Data.SqlClient.SqlParameter() = Nothing
        Return ExecuteNonQuery(connection, commandType, commandText, commandParameters)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns no resultset) against the specified SqlConnection using the provided parameters
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' Dim result as Integer = ExecuteNonQuery(conn, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal connection As System.Data.SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As System.Data.SqlClient.SqlParameter()) As Integer
        If (connection Is Nothing) Then Return 0
        'Create a command and prepare it for execution
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters, mustCloseConnection)
        'Finally, execute the command
        Dim retval As Integer = command.ExecuteNonQuery()
        'Detach the SqlParameters from the command object, so they can be used again
        command.Parameters.Clear()
        If (mustCloseConnection) Then
            connection.Close()
        End If
        Return retval
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns no resultset) against the specified SqlConnection
    ''' using the provided parameter values.  This method will query the database to discover the parameters for the
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Eg: Dim result As Integer = ExecuteNonQuery(conn, "PublishOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal connection As System.Data.SqlClient.SqlConnection, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As Integer
        If (connection Is Nothing) Then Return 0
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As System.Data.SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)
            'Call the overload that takes an array of SqlParameters
            Return ExecuteNonQuery(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteNonQuery(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns no resultset and takes no parameters) against the provided SqlTransaction. 
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' Dim result As Integer = ExecuteNonQuery(trans, CommandType.StoredProcedure, "PublishOrders")
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal transaction As System.Data.SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String) As Integer
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteNonQuery(transaction, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns no resultset) against the specified SqlTransaction using the provided parameters
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' Dim result as Integer = ExecuteNonQuery(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal transaction As System.Data.SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As System.Data.SqlClient.SqlParameter()) As Integer
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        'Create a command and prepare it for execution
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)
        'Finally, execute the command
        Dim retval As Integer = command.ExecuteNonQuery()
        'Detach the SqlParameters from the command object, so they can be used again
        command.Parameters.Clear()
        Return retval
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns no resultset) against the specified 
    ''' SqlTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Dim result As Integer = ExecuteNonQuery(conn, trans, "PublishOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteNonQuery(ByVal transaction As System.Data.SqlClient.SqlTransaction, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As Integer
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As System.Data.SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)

            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'Call the overload that takes an array of SqlParameters
            Return ExecuteNonQuery(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteNonQuery(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function

#End Region

#Region "-----> ExcuteDataReader <-----"

    ''' <summary>
    ''' This enum is used to indicate whether the connection was provided by the caller, or created by SqlHelper, so that
    ''' we can set the appropriate CommandBehavior when calling ExecuteReader()
    ''' </summary>
    ''' <remarks></remarks>
    Enum SqlConnectionOwnership
        ''' <summary>
        ''' Connection is owned and managed by SqlHelper
        ''' </summary>
        ''' <remarks></remarks>
        Internal
        ''' <summary>
        ''' Connection is owned and managed by the caller
        ''' </summary>
        ''' <remarks></remarks>
        External
    End Enum

    Public Shared Function ExcuteDataReader(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As SqlClient.SqlDataReader
        If (connection Is Nothing) Then Return Nothing
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters)
        Dim reader As SqlClient.SqlDataReader = command.ExecuteReader()
        command.Parameters.Clear()
        Return reader
    End Function

    Public Shared Function ExcuteDataReader(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal commandParameters As SqlClient.SqlParameterCollection) As SqlClient.SqlDataReader
        If (connection Is Nothing) Then Return Nothing
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters)
        Dim reader As SqlClient.SqlDataReader = command.ExecuteReader()
        command.Parameters.Clear()
        Return reader
    End Function

    ''' <summary>
    ''' Create and prepare a SqlCommand, and call ExecuteReader with the appropriate CommandBehavior
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection, on which to execute this command</param>
    ''' <param name="transaction">A valid SqlTransaction, or 'Nothing'</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParameters to be associated with the command or 'nothing' if no parameters are required</param>
    ''' <param name="connectionOwnership">Indicates whether the connection parameter was provided by the caller, or created by SqlHelper</param>
    ''' <returns>SqlDataReader containing the results of the command</returns>
    ''' <remarks>
    ''' If we created and opened the connection, we want the connection to be closed when the DataReader is closed.
    ''' If the caller provided the connection, we want to leave it to them to manage
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal connection As SqlClient.SqlConnection, ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal commandParameters As SqlClient.SqlParameter(), ByVal connectionOwnership As SqlConnectionOwnership) As SqlClient.SqlDataReader
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        Dim mustCloseConnection As Boolean = False
        'Create a command and prepare it for execution
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        Try
            PrepareCommand(command, connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)
            'Create a reader
            Dim dataReader As SqlClient.SqlDataReader
            'Call ExecuteReader with the appropriate CommandBehavior
            If (connectionOwnership = SqlConnectionOwnership.External) Then
                dataReader = command.ExecuteReader()
            Else
                dataReader = command.ExecuteReader(CommandBehavior.CloseConnection)
            End If
            'Detach the SqlParameters from the command object, so they can be used again
            'HACK: There is a problem here, the output parameter values are fletched 
            'when the reader is closed, so if the parameters are detached from the command then the SqlReader can´t set its values.
            ' When this happen, the parameters can´t be used again in other command.
            Dim canClear As Boolean = True
            For Each commandParameter As SqlClient.SqlParameter In command.Parameters
                If (commandParameter.Direction <> ParameterDirection.Input) Then
                    canClear = False
                End If
            Next
            If (canClear) Then
                command.Parameters.Clear()
            End If
            Return dataReader
        Catch ex As Exception
            If (mustCloseConnection) Then
                connection.Close()
            End If
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the database specified in the connection string
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim dr As SqlDataReader = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders")
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String) As SqlClient.SqlDataReader
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteReader(connectionString, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the database specified in the connection string using the provided parameters
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim dr As SqlDataReader = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As SqlClient.SqlDataReader
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        Dim connection As SqlClient.SqlConnection = Nothing
        Try
            connection = New SqlClient.SqlConnection(connectionString)
            connection.Open()
            'Call the private overload that takes an internally owned connection in place of the connection string
            Return ExecuteReader(connection, Nothing, commandType, commandText, commandParameters, SqlConnectionOwnership.Internal)
        Catch
            'If we fail to return the SqlDatReader, we need to close the connection ourselves
            If (connection IsNot Nothing) Then connection.Close()
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in 
    ''' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Eg: Dim dr As SqlDataReader = ExecuteReader(connString, "GetOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal connectionString As String, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As SqlClient.SqlDataReader
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)
            AssignParameterValues(commandParameters, parameterValues)
            Return ExecuteReader(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteReader(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlConnection
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim dr As SqlDataReader = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders")
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String) As SqlClient.SqlDataReader
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteReader(connection, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection using the provided parameters
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim dr As SqlDataReader = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As SqlClient.SqlDataReader
        'Pass through the call to the private overload using a null transaction value and an externally owned connection
        Return ExecuteReader(connection, Nothing, commandType, commandText, commandParameters, SqlConnectionOwnership.External)
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ''' using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Eg: Dim dr As SqlDataReader = ExecuteReader(conn, "GetOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As SqlClient.SqlDataReader
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            AssignParameterValues(commandParameters, parameterValues)
            Return ExecuteReader(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteReader(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlTransaction. 
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim dr As SqlDataReader = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders")
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String) As SqlClient.SqlDataReader
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteReader(transaction, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlTransaction using the provided parameters.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim dr as SqlDataReader = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As SqlClient.SqlDataReader
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        'Pass through to private overload, indicating that the connection is owned by the caller
        Return ExecuteReader(transaction.Connection, transaction, commandType, commandText, commandParameters, SqlConnectionOwnership.External)
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified
    ''' SqlTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Dim dr As SqlDataReader = ExecuteReader(trans, "GetOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteReader(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As SqlClient.SqlDataReader
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)
            AssignParameterValues(commandParameters, parameterValues)
            Return ExecuteReader(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteReader(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function

#End Region

#Region "-----> ExecuteScalar <-----"

    ''' <summary>
    ''' Execute a SqlCommand (that returns a 1x1 resultset) against the specified SqlConnection using the provided parameters
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' Eg: Dim orderCount As Integer = CType(ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24)),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As Object
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        'Create a command and prepare it for execution
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters, mustCloseConnection)

        'Execute the command & return the results
        Dim retval As Object = command.ExecuteScalar()

        'Detach the SqlParameters from the command object, so they can be used again
        command.Parameters.Clear()
        If (mustCloseConnection) Then
            connection.Close()
        End If
        Return retval
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a 1x1 resultset) against the database specified in the connection string using the provided parameters
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim orderCount As Integer = Ctype(ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24)),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As Object
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        'Create & open a SqlConnection, and dispose of it after we are done
        Using connection As SqlClient.SqlConnection = New SqlClient.SqlConnection(connectionString)
            connection.Open()
            'Call the overload that takes a connection in place of the connection string
            Return ExecuteScalar(connection, commandType, commandText, commandParameters)
        End Using
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a 1x1 resultset and takes no parameters) against the database specified in the connection string.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim orderCount As Integer = Ctype(ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount"),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String) As Object
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteScalar(connectionString, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the database specified in 
    ''' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Dim orderCount As Integer = CType(ExecuteScalar(connString, "GetOrderCount", 24, 36),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal connectionString As String, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As Object
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If (Not (parameterValues Is Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)
            'Call the overload that takes an array of SqlParameters
            Return ExecuteScalar(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteScalar(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a 1x1 resultset and takes no parameters) against the provided SqlConnection.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim orderCount As Integer = CType(ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount"),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String) As Object
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteScalar(connection, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the specified SqlConnection 
    ''' using the provided parameter values. This method will query the database to discover the parameters for the
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Dim orderCount As Integer = CType(ExecuteScalar(conn, "GetOrderCount", 24, 36),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As Object
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If (Not (parameterValues Is Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)
            'Call the overload that takes an array of SqlParameters
            Return ExecuteScalar(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteScalar(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a 1x1 resultset and takes no parameters) against the provided SqlTransaction.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim orderCount As Integer = CType(ExecuteScalar(trans, CommandType.StoredProcedure, "GetOrderCount"),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String) As Object
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteScalar(transaction, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a 1x1 resultset) against the specified SqlTransaction using the provided parameters
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim orderCount As Integer = CType(ExecuteScalar(trans, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24)),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As Object
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If (Not (transaction Is Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")

        'Create a command and prepare it for execution
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)

        'Execute the command & return the results
        Dim retval As Object = command.ExecuteScalar()
        'Detach the SqlParameters from the command object, so they can be used again
        command.Parameters.Clear()
        Return retval
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the specified
    ''' SqlTransaction using the provided parameter values. This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Dim orderCount As Integer = CType(ExecuteScalar(trans, "GetOrderCount", 24, 36),Integer)
    ''' </remarks>
    Public Shared Function ExecuteScalar(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As Object
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If (Not (transaction Is Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'Call the overload that takes an array of SqlParameters
            Return ExecuteScalar(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteScalar(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function

#End Region

#Region "-----> ExecuteXmlReader <-----"

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection using the provided parameters
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command using "FOR XML AUTO"</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An XmlReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim r as System.Xml.XmlReader = ExecuteXmlReader(conn, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteXmlReader(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As System.Xml.XmlReader
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        Dim mustCloseConnection As Boolean = False
        'Create a command and prepare it for execution
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()

        Try
            PrepareCommand(command, connection, Nothing, commandType, commandText, commandParameters, mustCloseConnection)
            'Create the DataAdapter & DataSet
            Dim retval As System.Xml.XmlReader = command.ExecuteXmlReader()

            'Detach the SqlParameters from the command object, so they can be used again
            command.Parameters.Clear()
            Return retval
        Catch
            If (mustCloseConnection) Then
                connection.Close()
            End If
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlConnection. 
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command using "FOR XML AUTO"</param>
    ''' <returns>An XmlReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim r as System.Xml.XmlReader = ExecuteXmlReader(conn, CommandType.StoredProcedure, "GetOrders")
    ''' </remarks>
    Public Shared Function ExecuteXmlReader(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String) As System.Xml.XmlReader
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteXmlReader(connection, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ''' using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure using "FOR XML AUTO"</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>An XmlReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Dim r as System.Xml.XmlReader = ExecuteXmlReader(conn, "GetOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteXmlReader(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As System.Xml.XmlReader
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If (Not (parameterValues Is Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)
            'Call the overload that takes an array of SqlParameters
            Return ExecuteXmlReader(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteXmlReader(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlTransaction using the provided parameters
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command using "FOR XML AUTO"</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <returns>An XmlReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim r as System.Xml.XmlReader = ExecuteXmlReader(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Function ExecuteXmlReader(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal ParamArray commandParameters As SqlClient.SqlParameter()) As System.Xml.XmlReader
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If (Not (transaction Is Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        'Create a command and prepare it for execution
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)
        'Create the DataAdapter & DataSet
        Dim retval As System.Xml.XmlReader = command.ExecuteXmlReader()
        'Detach the SqlParameters from the command object, so they can be used again
        command.Parameters.Clear()
        Return retval
    End Function

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlTransaction
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command using "FOR XML AUTO"</param>
    ''' <returns>An XmlReader containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' Dim r as System.Xml.XmlReader = ExecuteXmlReader(trans, CommandType.StoredProcedure, "GetOrders")
    ''' </remarks>
    Public Shared Function ExecuteXmlReader(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String) As System.Xml.XmlReader
        'Pass through the call providing null for the set of SqlParameters
        Return ExecuteXmlReader(transaction, commandType, commandText, Nothing)
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified 
    ''' SqlTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' Dim r as System.Xml.XmlReader = ExecuteXmlReader(trans, "GetOrders", 24, 36)
    ''' </remarks>
    Public Shared Function ExecuteXmlReader(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal ParamArray parameterValues As Object()) As System.Xml.XmlReader
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If (Not (transaction Is Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)

            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'Call the overload that takes an array of SqlParameters
            Return ExecuteXmlReader(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            Return ExecuteXmlReader(transaction, CommandType.StoredProcedure, spName)
        End If

    End Function

#End Region

#Region "-----> Fill Dataset <-----"

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the database specified in the connection string
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">
    ''' This array will be used to create table mappings allowing the DataTables to be referenced
    ''' by a user defined name (probably the actual table name)
    ''' </param>
    ''' <remarks>
    ''' FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"})
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String, ByVal dataSet As DataSet, ByVal tableNames As String())
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If (dataSet Is Nothing) Then Throw New ArgumentNullException("dataSet")
        'Create & open a SqlConnection, and dispose of it after we are done
        Using connection As SqlClient.SqlConnection = New SqlClient.SqlConnection(connectionString)
            connection.Open()
            'Call the overload that takes a connection in place of the connection string
            FillDataset(connection, commandType, commandText, dataSet, tableNames)
        End Using

    End Sub

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the database specified in the connection string using the provided parameters.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="dataSet">An array of SqlParamters used to execute the command</param>
    ''' <param name="tableNames">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="commandParameters">
    ''' This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)
    ''' </param>
    ''' <remarks>
    ''' FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"}, new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal connectionString As String, ByVal commandType As CommandType, ByVal commandText As String, ByVal dataSet As DataSet, ByVal tableNames As String(), ByVal ParamArray commandParameters As SqlClient.SqlParameter())
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If (dataSet Is Nothing) Then Throw New ArgumentNullException("dataSet")
        'Create & open a SqlConnection, and dispose of it after we are done
        Using connection As SqlClient.SqlConnection = New SqlClient.SqlConnection(connectionString)
            connection.Open()
            'Call the overload that takes a connection in place of the connection string
            FillDataset(connection, commandType, commandText, dataSet, tableNames, commandParameters)
        End Using
    End Sub

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in 
    ''' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter.
    ''' Eg: FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"}, 24)
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal connectionString As String, ByVal spName As String, ByVal dataSet As DataSet, ByVal tableNames As String(), ByVal ParamArray parameterValues As Object())
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If (dataSet Is Nothing) Then Throw New ArgumentNullException("dataSet")
        'Create & open a SqlConnection, and dispose of it after we are done
        Using connection As SqlClient.SqlConnection = New SqlClient.SqlConnection(connectionString)
            connection.Open()
            'Call the overload that takes a connection in place of the connection string
            FillDataset(connection, spName, dataSet, tableNames, parameterValues)
        End Using
    End Sub

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlConnection
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <remarks>
    ''' Eg: FillDataset(conn, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"})
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal dataSet As DataSet, ByVal tableNames As String())
        FillDataset(connection, commandType, commandText, dataSet, tableNames, Nothing)
    End Sub

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection using the provided parameters
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <remarks>
    ''' FillDataset(conn, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"}, new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal connection As SqlClient.SqlConnection, ByVal commandType As CommandType, ByVal commandText As String, ByVal dataSet As DataSet, ByVal tableNames As String(), ByVal ParamArray commandParameters As SqlClient.SqlParameter())
        FillDataset(connection, Nothing, commandType, commandText, dataSet, tableNames, commandParameters)
    End Sub

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ''' using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' FillDataset(conn, "GetOrders", ds, new string() {"orders"}, 24, 36)
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal dataSet As DataSet, ByVal tableNames As String(), ByVal ParamArray parameterValues As Object())
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If (dataSet Is Nothing) Then Throw New ArgumentNullException("dataSet")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            ' Call the overload that takes an array of SqlParameters
            FillDataset(connection, CommandType.StoredProcedure, spName, dataSet, tableNames, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            FillDataset(connection, CommandType.StoredProcedure, spName, dataSet, tableNames)
        End If
    End Sub

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlTransaction. 
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <remarks>
    ''' FillDataset(trans, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"})
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal dataSet As DataSet, ByVal tableNames As String())
        FillDataset(transaction, commandType, commandText, dataSet, tableNames, Nothing)
    End Sub

    ''' <summary>
    ''' Execute a SqlCommand (that returns a resultset) against the specified SqlTransaction using the provided parameters
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <remarks>
    ''' Eg: FillDataset(trans, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"}, new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal dataSet As DataSet, ByVal tableNames As String(), ByVal ParamArray commandParameters As SqlClient.SqlParameter())
        FillDataset(transaction.Connection, transaction, commandType, commandText, dataSet, tableNames, commandParameters)
    End Sub


    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified
    ''' SqlTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
    ''' <remarks>
    ''' This method provides no access to output parameters or the stored procedure's return value parameter
    ''' FillDataset(trans, "GetOrders", ds, new string() {"orders"}, 24, 36)
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal dataSet As DataSet, ByVal tableNames As String(), ByVal ParamArray parameterValues As Object())
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If (dataSet Is Nothing) Then Throw New ArgumentNullException("dataSet")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If we receive parameter values, we need to figure out where they go
        If ((parameterValues IsNot Nothing) And (parameterValues.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)

            'Assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'Call the overload that takes an array of SqlParameters
            FillDataset(transaction, CommandType.StoredProcedure, spName, dataSet, tableNames, commandParameters)
        Else
            'Otherwise we can just call the SP without params
            FillDataset(transaction, CommandType.StoredProcedure, spName, dataSet, tableNames)
        End If
    End Sub

    ''' <summary>
    ''' Private helper method that execute a SqlCommand (that returns a resultset) against the specified SqlTransaction and SqlConnection using the provided parameters
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection</param>
    ''' <param name="transaction">A valid SqlTransaction</param>
    ''' <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
    ''' <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced by a user defined name (probably the actual table name)</param>
    ''' <param name="commandParameters">An array of SqlParamters used to execute the command</param>
    ''' <remarks>
    ''' Eg: FillDataset(conn, trans, CommandType.StoredProcedure, "GetOrders", ds, new string() {"orders"}, new SqlParameter("@prodid", 24))
    ''' </remarks>
    Public Shared Sub FillDataset(ByVal connection As SqlClient.SqlConnection, ByVal transaction As SqlClient.SqlTransaction, ByVal commandType As CommandType, ByVal commandText As String, ByVal dataSet As DataSet, ByVal tableNames As String(), ByVal ParamArray commandParameters As SqlClient.SqlParameter())
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If (dataSet Is Nothing) Then Throw New ArgumentNullException("dataSet")
        'Create a command and prepare it for execution
        Dim command As SqlClient.SqlCommand = New SqlClient.SqlCommand()
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(command, connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)
        'Create the DataAdapter & DataSet
        Using mydap As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(command)
            ' Add the table mappings specified by the user
            If ((tableNames IsNot Nothing) And (tableNames.Length > 0)) Then
                Dim tableName As String = "Table"
                For i As Integer = 0 To tableNames.Length - 1
                    If (tableNames(i) = Nothing Or tableNames(i).Length = 0) Then Throw New ArgumentException("The tableNames parameter must contain a list of tables, a value was provided as null or empty string.", "tableNames")
                    mydap.TableMappings.Add(tableName, tableNames(i))
                    tableName += (i + 1).ToString()
                Next
            End If
            'Fill the DataSet using default values for DataTable names, etc
            mydap.Fill(dataSet)
            'Detach the SqlParameters from the command object, so they can be used again
            command.Parameters.Clear()
        End Using
        If (mustCloseConnection) Then
            connection.Close()
        End If
    End Sub

#End Region

#Region "-----> UpdateDataset <-----"

    ''' <summary>
    ''' Executes the respective command for each inserted, updated, or deleted row in the DataSet.
    ''' </summary>
    ''' <param name="insertCommand">A valid transact-SQL statement or stored procedure to insert new records into the data source</param>
    ''' <param name="deleteCommand">A valid transact-SQL statement or stored procedure to delete records from the data source</param>
    ''' <param name="updateCommand">A valid transact-SQL statement or stored procedure used to update records in the data source</param>
    ''' <param name="dataSet">The DataSet used to update the data source</param>
    ''' <param name="tableName">The DataTable used to update the data source</param>
    ''' <remarks>
    ''' UpdateDataset(conn, insertCommand, deleteCommand, updateCommand, dataSet, "Order")
    ''' </remarks>
    Public Shared Sub UpdateDataset(ByVal insertCommand As SqlClient.SqlCommand, ByVal deleteCommand As SqlClient.SqlCommand, ByVal updateCommand As SqlClient.SqlCommand, ByVal dataSet As DataSet, ByVal tableName As String)
        If (insertCommand Is Nothing) Then Throw New ArgumentNullException("InsertCommand")
        If (deleteCommand Is Nothing) Then Throw New ArgumentNullException("DeleteCommand")
        If (updateCommand Is Nothing) Then Throw New ArgumentNullException("UpdateCommand")
        If ((tableName = Nothing) Or (tableName.Length = 0)) Then Throw New ArgumentNullException("TableName")
        'Create a SqlDataAdapter, and dispose of it after we are done
        Using mydap As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter()
            'Set the data adapter commands
            mydap.UpdateCommand = updateCommand
            mydap.InsertCommand = insertCommand
            mydap.DeleteCommand = deleteCommand
            'Update the dataset changes in the data source
            mydap.Update(dataSet, tableName)
            'Commit all the changes made to the DataSet
            dataSet.AcceptChanges()
        End Using
    End Sub

#End Region

#Region "-----> ExecuteNonQueryTypedParams <-----"
    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns no resultset) against the database specified in 
    ''' the connection string using the dataRow column values as the stored procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on row values.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteNonQueryTypedParams(ByVal connectionString As String, ByVal spName As String, ByVal dataRow As DataRow) As Integer
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns no resultset) against the specified SqlConnection 
    ''' using the dataRow column values as the stored procedure's parameters values.  
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on row values.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteNonQueryTypedParams(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal dataRow As DataRow) As Integer
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)

            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)

            Return SoftSqlHelper.ExecuteNonQuery(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteNonQuery(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns no resultset) against the specified
    ''' SqlTransaction using the dataRow column values as the stored procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on row values.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>An int representing the number of rows affected by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteNonQueryTypedParams(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal dataRow As DataRow) As Integer
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)

            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteNonQuery(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteNonQuery(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function
#End Region

#Region "-----> ExecuteDatasetTypedParams <-----"

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in
    ''' the connection string using the dataRow column values as the stored procedure's parameters values
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on row values.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteDatasetTypedParams(ByVal connectionString As String, ByVal spName As String, ByVal dataRow As DataRow) As DataSet
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)

            Return SoftSqlHelper.ExecuteDataset(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteDataset(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ''' using the dataRow column values as the store procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on row values.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteDatasetTypedParams(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal dataRow As DataRow) As DataSet
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteDataset(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteDataset(connection, CommandType.StoredProcedure, spName)
        End If

    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlTransaction 
    ''' using the dataRow column values as the stored procedure's parameters values
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on row values.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>A dataset containing the resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteDatasetTypedParams(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal dataRow As DataRow) As DataSet
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteDataset(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteDataset(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function

#End Region

#Region "-----> ExecuteReaderTypedParams <-----"

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in 
    ''' the connection string using the dataRow column values as the stored procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteReaderTypedParams(ByVal connectionString As String, ByVal spName As String, ByVal dataRow As DataRow) As SqlClient.SqlDataReader
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteReader(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteReader(connectionString, CommandType.StoredProcedure, spName)
        End If

    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ''' using the dataRow column values as the stored procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteReaderTypedParams(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal dataRow As DataRow) As SqlClient.SqlDataReader
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteReader(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteReader(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlTransaction 
    ''' using the dataRow column values as the stored procedure's parameters values
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>A SqlDataReader containing the resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteReaderTypedParams(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal dataRow As DataRow) As SqlClient.SqlDataReader
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteReader(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteReader(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function

#End Region

#Region "-----> ExecuteScalarTypedParams <-----"

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the database specified in 
    ''' the connection string using the dataRow column values as the stored procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteScalarTypedParams(ByVal connectionString As String, ByVal spName As String, ByVal dataRow As DataRow) As Object
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)

            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteScalar(connectionString, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteScalar(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the specified SqlConnection 
    ''' using the dataRow column values as the stored procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="connection">A valid SqlConnection object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the comman</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteScalarTypedParams(ByVal connection As SqlClient.SqlConnection, ByVal spName As String, ByVal dataRow As DataRow) As Object
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(connection, spName)
            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteScalar(connection, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteScalar(connection, CommandType.StoredProcedure, spName)
        End If
    End Function

    ''' <summary>
    ''' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the specified SqlTransaction
    ''' using the dataRow column values as the stored procedure's parameters values.
    ''' This method will query the database to discover the parameters for the 
    ''' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
    ''' </summary>
    ''' <param name="transaction">A valid SqlTransaction objec</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="dataRow">The dataRow used to hold the stored procedure's parameter values</param>
    ''' <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteScalarTypedParams(ByVal transaction As SqlClient.SqlTransaction, ByVal spName As String, ByVal dataRow As DataRow) As Object
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If ((transaction IsNot Nothing) And (transaction.Connection Is Nothing)) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
        If ((spName = Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        'If the row has values, the store procedure parameters must be initialized
        If ((dataRow IsNot Nothing) And (dataRow.ItemArray.Length > 0)) Then
            'Pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            Dim commandParameters As SqlClient.SqlParameter() = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection, spName)

            'Set the parameters values
            AssignParameterValues(commandParameters, dataRow)
            Return SoftSqlHelper.ExecuteScalar(transaction, CommandType.StoredProcedure, spName, commandParameters)
        Else
            Return SoftSqlHelper.ExecuteScalar(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function

    

#End Region

    Public Shared Function GetString(ByVal strSQL As String, ByVal defVal As String) As String
        Using _connnection As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            If (_connnection.State.Equals(Nothing) Or (_connnection.State = ConnectionState.Closed)) Then
                _connnection.Open()
            End If
            Dim mySqlCommand As System.Data.SqlClient.SqlCommand = New SqlClient.SqlCommand(strSQL, _connnection)
            Try
                Dim valueResult As Object = mySqlCommand.ExecuteScalar()
                Return Convert.ToString(valueResult)
            Catch
                Return defVal
            Finally
                DbCommon.CloseConnection(_connnection)
            End Try

        End Using
    End Function

    Public Shared Function GetNumber(ByVal strSQL As String, ByVal defVal As Integer) As Integer
        Using _connnection As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            If (_connnection.State.Equals(Nothing) Or (_connnection.State = ConnectionState.Closed)) Then
                _connnection.Open()
            End If
            Dim mySqlCommand As System.Data.SqlClient.SqlCommand = New SqlClient.SqlCommand(strSQL, _connnection)
            Try
                Dim valueResult As Object = mySqlCommand.ExecuteScalar()
                Return Convert.ToInt32(valueResult)
            Catch
                Return defVal
            Finally
                DbCommon.CloseConnection(_connnection)
            End Try

        End Using
    End Function

    Public Shared Function GetNumberDouble(ByVal strSQL As String, ByVal defVal As Double) As Double
        Using _connnection As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            If (_connnection.State.Equals(Nothing) Or (_connnection.State = ConnectionState.Closed)) Then
                _connnection.Open()
            End If
            Dim mySqlCommand As System.Data.SqlClient.SqlCommand = New SqlClient.SqlCommand(strSQL, _connnection)
            Try
                Dim valueResult As Object = mySqlCommand.ExecuteScalar()
                Return Convert.ToDouble(valueResult)
            Catch
                Return defVal
            Finally
                DbCommon.CloseConnection(_connnection)
            End Try

        End Using
    End Function

    Public Shared Function GetNumberLong(ByVal strSQL As String, ByVal defVal As Long) As Long
        Using _connnection As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            If (_connnection.State.Equals(Nothing) Or (_connnection.State = ConnectionState.Closed)) Then
                _connnection.Open()
            End If
            Dim mySqlCommand As System.Data.SqlClient.SqlCommand = New SqlClient.SqlCommand(strSQL, _connnection)
            Try
                Dim valueResult As Object = mySqlCommand.ExecuteScalar()
                Return Convert.ToInt64(valueResult)
            Catch
                Return defVal
            Finally
                DbCommon.CloseConnection(_connnection)
            End Try

        End Using
    End Function
End Class

''' <summary>
''' SqlHelperParameterCache provides functions to leverage a static cache of procedure parameters, and the
''' ability to discover parameters for stored procedures at run-time.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class SqlHelperParameterCache

#Region "---> Private methods, variables, and constructors <---"

    'Since this class provides only static methods, make the default constructor private to prevent instances from being created with "new SqlHelperParameterCache()"
    Private Sub New()

    End Sub

    Private Shared paramCache As System.Collections.Hashtable = System.Collections.Hashtable.Synchronized(New Hashtable())

    ''' <summary>
    ''' Resolve at run time the appropriate set of SqlParameters for a stored procedure
    ''' </summary>
    ''' <param name="_connection">A valid SqlConnection object</param>
    ''' <param name="_spName">The name of the stored procedure</param>
    ''' <param name="includeReturnValueParameter">Whether or not to include their return value parameter</param>
    ''' <returns>The parameter array discovered.</returns>
    ''' <remarks></remarks>
    Private Shared Function DiscoverSpParameterSet(ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal _spName As String, ByVal includeReturnValueParameter As Boolean) As System.Data.SqlClient.SqlParameter()
        If (_connection Is Nothing) Then Throw New ArgumentNullException("_connection")
        If ((_spName Is Nothing) Or (_spName.Length = 0)) Then Throw New ArgumentNullException("_spName")

        Dim _command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(_spName, _connection)
        _command.CommandType = CommandType.StoredProcedure
        _connection.Open()
        System.Data.SqlClient.SqlCommandBuilder.DeriveParameters(_command)
        _connection.Close()
        If (Not (includeReturnValueParameter)) Then
            _command.Parameters.RemoveAt(0)
        End If
        Dim _discoveredParameters(_command.Parameters.Count) As System.Data.SqlClient.SqlParameter
        _command.Parameters.CopyTo(_discoveredParameters, 0)
        'Init the parameters with a DBNull value
        For Each _discoveredParameter As System.Data.SqlClient.SqlParameter In _discoveredParameters
            _discoveredParameter.Value = DBNull.Value
        Next
        Return _discoveredParameters
    End Function

    ''' <summary>
    ''' Deep copy of cached SqlParameter array
    ''' </summary>
    ''' <param name="originalParameters"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CloneParameters(ByVal originalParameters As System.Data.SqlClient.SqlParameter()) As System.Data.SqlClient.SqlParameter()
        Dim clonedParameters(originalParameters.Length) As System.Data.SqlClient.SqlParameter
        Dim j As Integer = originalParameters.Length - 1
        For i As Integer = 0 To j
            clonedParameters(i) = CType(CType(originalParameters(i), ICloneable).Clone(), System.Data.SqlClient.SqlParameter)
        Next
        Return clonedParameters
    End Function

#End Region

#Region "---> Caching functions <---"

    ''' <summary>
    ''' Add parameter array to the cache
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <param name="commandParameters">An array of SqlParamters to be cached</param>
    ''' <remarks></remarks>
    Public Shared Sub CacheParameterSet(ByVal connectionString As String, ByVal commandText As String, ByVal ParamArray commandParameters As System.Data.SqlClient.SqlParameter())
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((commandText Is Nothing) Or (commandText.Length = 0)) Then Throw New ArgumentNullException("commandText")
        'string hashKey = connectionString + ":" + commandText
        Dim hashKey As String = connectionString + ":" + commandText
        paramCache(hashKey) = commandParameters
    End Sub

    ''' <summary>
    ''' Retrieve a parameter array from the cache
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="commandText">The stored procedure name or T-SQL command</param>
    ''' <returns>An array of SqlParamters</returns>
    ''' <remarks></remarks>
    Public Shared Function GetCachedParameterSet(ByVal connectionString As String, ByVal commandText As String) As System.Data.SqlClient.SqlParameter()
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((commandText Is Nothing) Or (commandText.Length = 0)) Then Throw New ArgumentNullException("commandText")
        Dim hashKey As String = connectionString + ":" + commandText
        Dim cachedParameters As System.Data.SqlClient.SqlParameter() = CType(paramCache(hashKey), System.Data.SqlClient.SqlParameter())
        If (cachedParameters Is Nothing) Then
            Return Nothing
        Else
            Return CloneParameters(cachedParameters)
        End If
    End Function

#End Region

#Region "---> Parameter Discovery Functions <---"
    ''' <summary>
    ''' Retrieves the set of SqlParameters appropriate for the stored procedure
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <returns>An array of SqlParameters</returns>
    ''' <remarks>This method will query the database for this information, and then store it in a cache for future requests</remarks>
    Public Shared Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String) As System.Data.SqlClient.SqlParameter()
        Return GetSpParameterSet(connectionString, spName, False)
    End Function

    ''' <summary>
    ''' Retrieves the set of SqlParameters appropriate for the stored procedure
    ''' </summary>
    ''' <param name="connectionString">A valid connection string for a SqlConnection</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="includeReturnValueParameter">A bool value indicating whether the return value parameter should be included in the results</param>
    ''' <returns>An array of SqlParameters</returns>
    ''' <remarks>This method will query the database for this information, and then store it in a cache for future requests</remarks>
    Public Shared Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String, ByVal includeReturnValueParameter As Boolean) As System.Data.SqlClient.SqlParameter()
        If ((connectionString = Nothing) Or (connectionString.Length = 0)) Then Throw New ArgumentNullException("connectionString")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        Using _connection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(connectionString)
            Return GetSpParameterSetInternal(_connection, spName, includeReturnValueParameter)
        End Using
    End Function

    ''' <summary>
    ''' Retrieves the set of SqlParameters appropriate for the stored procedure
    ''' </summary>
    ''' <param name="_connection">A valid SqlConnection object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <returns>An array of SqlParameters</returns>
    ''' <remarks>This method will query the database for this information, and then store it in a cache for future requests.</remarks>
    Friend Shared Function GetSpParameterSet(ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal spName As String) As System.Data.SqlClient.SqlParameter()
        Return GetSpParameterSet(_connection, spName, False)
    End Function

    ''' <summary>
    ''' Retrieves the set of SqlParameters appropriate for the stored procedure
    ''' </summary>
    ''' <param name="_connection">A valid SqlConnection object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="includeReturnValueParameter">A bool value indicating whether the return value parameter should be included in the results</param>
    ''' <returns>An array of SqlParameters</returns>
    ''' <remarks>This method will query the database for this information, and then store it in a cache for future requests.</remarks>
    Friend Shared Function GetSpParameterSet(ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal spName As String, ByVal includeReturnValueParameter As Boolean) As System.Data.SqlClient.SqlParameter()
        If (_connection Is Nothing) Then Throw New ArgumentNullException("_connection")
        Using clonedConnection As System.Data.SqlClient.SqlConnection = CType(CType(_connection, ICloneable).Clone(), System.Data.SqlClient.SqlConnection)
            Return GetSpParameterSetInternal(clonedConnection, spName, includeReturnValueParameter)
        End Using
    End Function

    ''' <summary>
    ''' Retrieves the set of SqlParameters appropriate for the stored procedure
    ''' </summary>
    ''' <param name="_connection">A valid SqlConnection object</param>
    ''' <param name="spName">The name of the stored procedure</param>
    ''' <param name="includeReturnValueParameter">A bool value indicating whether the return value parameter should be included in the results</param>
    ''' <returns>An array of SqlParameters</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSpParameterSetInternal(ByVal _connection As System.Data.SqlClient.SqlConnection, ByVal spName As String, ByVal includeReturnValueParameter As Boolean) As System.Data.SqlClient.SqlParameter()
        If (_connection Is Nothing) Then Throw New ArgumentNullException("_connection")
        If ((spName Is Nothing) Or (spName.Length = 0)) Then Throw New ArgumentNullException("spName")
        Dim hashKey As String = _connection.ConnectionString + ":" + spName + IIf(includeReturnValueParameter, ":include ReturnValue Parameter", "")
        Dim cachedParameters As System.Data.SqlClient.SqlParameter()
        cachedParameters = CType(paramCache(hashKey), System.Data.SqlClient.SqlParameter())
        If (cachedParameters Is Nothing) Then
            'MySqlParameter[] spParameters = DiscoverSpParameterSet(connection, spName, includeReturnValueParameter)
            Dim spParameters As System.Data.SqlClient.SqlParameter() = DiscoverSpParameterSet(_connection, spName, includeReturnValueParameter)
            paramCache(hashKey) = spParameters
            cachedParameters = spParameters
        End If
        Return CloneParameters(cachedParameters)
    End Function
#End Region

End Class