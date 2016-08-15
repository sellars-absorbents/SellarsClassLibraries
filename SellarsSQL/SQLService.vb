'//--------------------------------------------------------------------------------------------------------
'// Declare all the NameSpaces to be referenced
'//--------------------------------------------------------------------------------------------------------
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Collections

'//------------------------------------------------------------------------------------------------------------
'// Filename:	SQLService.vb
'// Author:		Stratecon, Inc.
'// Date:		09/11/2002
'// Purpose:	The DataLayer component is an interface between the business services tier and 
'// 			a data source. It can execute stored procedures or Sql statements and it can 
'//			    return rowsets, if requested.
'//------------------------------------------------------------------------------------------------------------

'//--------------------------------------------------------------------------------------------------------------
'// Public Class
'// Purpose:	The SqlService class interacts with a datasource using ADO.Net's 
'//				SqlClient interface.
'//--------------------------------------------------------------------------------------------------------------
Public Class SqlService

    '//----------------------------------------------------------------------------------------------------
    '// Declare Class level variables
    '//----------------------------------------------------------------------------------------------------
    Private m_sUsername As String = ""          '//--- The Database login User ID
    Private m_sPassword As String = ""          '//--- The Database login Password
    Private m_sServer As String = ""            '//--- The SQL Server instance
    Private m_sDatabase As String = ""          '//--- The Database name
    Private m_sConnectionString As String = ""  '//--- The Database connection string
    Private m_iErrorCode As Integer = 0         '//--- The database process error code
    Private m_sErrorDescription As String = ""  '//--- The database process error description

    '//--- The array used to store the parameters to a stored procedure
    Private m_oParmList As ArrayList = New ArrayList

    '//----------------------------------------------------------------------------------------------------
    '// Class Constructor (zero arguments)
    '// Overloaded:	Yes
    '//----------------------------------------------------------------------------------------------------
    Sub New()
    End Sub

    '//----------------------------------------------------------------------------------------------------
    '// Class Constructor (with the entire connection string as the argument)
    '// Overloaded:	Yes
    '//----------------------------------------------------------------------------------------------------
    Sub New(ByVal sConnectionString As String)
        m_sConnectionString = sConnectionString
    End Sub

    '//----------------------------------------------------------------------------------------------------
    '// Class Constructor (with connection string arguments)
    '// Overloaded:	Yes
    '//----------------------------------------------------------------------------------------------------
    Sub New(ByVal sServer As String, ByVal sDatabase As String, ByVal sUsername As String, ByVal sPassword As String)
        Server = sServer
        Database = sDatabase
        Username = sUsername
        Password = sPassword
    End Sub

    '//----------------------------------------------------------------------------------------------------
    '// Public R/W Property 
    '// Purpose:		Exposes the connection string
    '//----------------------------------------------------------------------------------------------------
    Public Property ConnectionString() As String
        Get
            Return m_sConnectionString
        End Get
        Set(ByVal Value As String)
            m_sConnectionString = Value
        End Set
    End Property

    '//----------------------------------------------------------------------------------------------------
    '// Public R/W Property 
    '// Purpose:		Exposes the error code value
    '//----------------------------------------------------------------------------------------------------
    Public Property ErrorCode() As Integer
        Get
            Return m_iErrorCode
        End Get
        Set(ByVal Value As Integer)
            m_iErrorCode = Value
        End Set
    End Property

    '//----------------------------------------------------------------------------------------------------
    '// Public R/W Property 
    '// Purpose:		Exposes the error description
    '//----------------------------------------------------------------------------------------------------
    Public Property ErrorDescription() As String
        Get
            Return m_sErrorDescription
        End Get
        Set(ByVal Value As String)
            m_sErrorDescription = Value
        End Set
    End Property

    '//----------------------------------------------------------------------------------------------------
    '// Public R/W Property 
    '// Purpose:		Exposes the User ID 
    '//----------------------------------------------------------------------------------------------------
    Public Property Username() As String
        Get
            Return m_sUsername
        End Get
        Set(ByVal Value As String)
            m_sUsername = Value
        End Set
    End Property

    '//------------------------------------------------------------------------------------------------------------
    '// Public R/W Property 
    '// Purpose:    Exposes the Password
    '//------------------------------------------------------------------------------------------------------------
    Public Property Password() As String
        Get
            Return m_sPassword
        End Get
        Set(ByVal Value As String)
            m_sPassword = Value
        End Set
    End Property

    '//------------------------------------------------------------------------------------------------------------
    '// Public R/W Property 
    '// Purpose:		Exposes the SQL Server
    '//------------------------------------------------------------------------------------------------------------
    Public Property Server() As String
        Get
            Return m_sServer
        End Get
        Set(ByVal Value As String)
            m_sServer = Value
        End Set
    End Property

    '//------------------------------------------------------------------------------------------------------------
    '// Public R/W Property
    '// Purpose:		Exposes the Database name
    '//------------------------------------------------------------------------------------------------------------
    Public Property Database() As String
        Get
            Return m_sDatabase
        End Get
        Set(ByVal Value As String)
            m_sDatabase = Value
        End Set
    End Property

    '//------------------------------------------------------------------------------------------------------------
    '// Public Method
    '// Overloaded:		Yes
    '// Return Value:	DataSet
    '// Purpose:		Executes a SQL statement.
    '//------------------------------------------------------------------------------------------------------------
    Public Overloads Function RunSql(ByVal sSql As String, ByVal sTableName As String) As DataSet

        Dim oCmd As SqlCommand = New SqlCommand             '//--- Create a new SqlCommand
        Dim oCn As SqlConnection = Nothing                  '//--- Declare the SqlConnection
        Dim oDa As SqlDataAdapter = New SqlDataAdapter      '//--- Create a new SqlDataAdapter
        Dim oDs As DataSet = New DataSet                    '//--- Create a new DataSet

        '//--- Prepare connection to the database
        oCn = Connect()

        With oCmd
            '//--- Set the CommandText, ActiveConnection and the Command Type 
            '//--- for the SqlCommand Object
            .Connection = oCn
            .CommandText = sSql
            .CommandType = CommandType.Text
        End With

        With oDa
            '//--- Assign the SqlCommand Object to the
            '//--- Select command of the SqlDataAdapter and
            .SelectCommand = oCmd

            '//--- Execute the Sql Statement and fill the dataset
            .Fill(oDs, sTableName)
        End With

        '//--- Disconnect from the database
        Disconnect(oCn)
        oCn = Nothing

        '//--- Return the DataSet
        Return oDs

    End Function

    '//------------------------------------------------------------------------------------------------------------
    '// Public Method
    '// Overloaded:		Yes
    '// Return Value:	None
    '// Purpose:		Executes a SQL statement and returns nothing.
    '//------------------------------------------------------------------------------------------------------------
    Public Overloads Sub RunSql(ByVal sSql As String)

        Dim oCmd As SqlCommand = New SqlCommand             '//--- Create a new SqlCommand
        Dim oCn As SqlConnection = Nothing                  '//--- Declare the SqlConnection

        '//--- Prepare connection to the database
        oCn = Connect()

        With oCmd
            '//--- Execute the Sql Satement
            .CommandText = sSql
            .Connection = oCn
            .CommandType = CommandType.Text
            .ExecuteNonQuery()
        End With

        '//--- Disconnect from the database
        Disconnect(oCn)
        oCn = Nothing
    End Sub

    '//------------------------------------------------------------------------------------------------------------
    '// Public Method
    '// Overloaded:		Yes
    '// Return Value:	DataSet
    '// Purpose:		Executes a stored procedure.
    '//------------------------------------------------------------------------------------------------------------
    Public Function RunProcDataSet(ByVal sProcName As String, ByVal sTableName As String) As DataSet

        Dim oCmd As SqlCommand = New SqlCommand             '//--- Create a new SqlCommand
        Dim oCn As SqlConnection = Nothing                  '//--- Declare the SqlConnection
        Dim oDA As SqlDataAdapter = New SqlDataAdapter      '//--- Create a new SqlDataAdapter
        Dim oDs As DataSet = New DataSet                    '//--- Create a new DataSet
        Dim oSqlParameter As SqlParameter = Nothing         '//--- Declare a SqlParameter
        Dim oP As Parameter = Nothing                       '//--- Declare a Parameter
        '//--- Get an enumerator for the parameter array list
        Dim oEnumerator As IEnumerator = m_oParmList.GetEnumerator()

        '//--- Prepare connection to the database
        oCn = Connect()

        With oCmd
            '//--- Set the CommandText, ActiveConnection and the Command Type for the 
            '//--- SqlCommand Object
            .Connection = oCn
            .CommandText = sProcName
            .CommandType = CommandType.StoredProcedure
        End With

        '//--- Add the parameters to the command
        AddParameters(oEnumerator, oCmd)

        ' Initialize the error code value
        ErrorCode = 0

        Try
            With oDA
                '//--- Assign the SqlCommand Object to the
                '//--- Select command of the SqlDataAdapter and
                .SelectCommand = oCmd

                '//--- Execute the Stored Procedure and fill the dataset
                .Fill(oDs, sTableName)
            End With
        Catch ex As Exception
            ' Capture the error codes
            ErrorCode = Err.Number
            ErrorDescription = ex.Message
        Finally

            '//--- Disconnect from the database
            Disconnect(oCn)
            oCn = Nothing
        End Try

        '//--- Get the output parameter values if there were no errors
        If ErrorCode = 0 Then
            GetOutputParameterValues(oCmd)
        End If

        '//--- Return the DataSet
        Return oDs

    End Function

    '//------------------------------------------------------------------------------------------------------------
    '// Public Method
    '// Overloaded:		Yes
    '// Return Value:	None
    '// Purpose:		Executes a stored procedure.
    '//------------------------------------------------------------------------------------------------------------
    Public Overloads Sub RunProc(ByVal sProcName As String)

        Dim oCmd As SqlCommand = New SqlCommand             '//--- Create a new SqlCommand
        Dim oCn As SqlConnection = Nothing                  '//--- Declare the SqlConnection
        Dim oSqlParameter As SqlParameter = Nothing         '//--- Declare a SqlParameter
        Dim oP As Parameter = Nothing                       '//--- Declare a Parameter
        '//--- Get an enumerator for the parameter array list
        Dim oEnumerator As IEnumerator = m_oParmList.GetEnumerator()

        '//--- Prepare connection to the database
        oCn = Connect()

        With oCmd
            '//--- Set the CommandText, ActiveConnection and the Command Type for the 
            '//--- SqlCommand Object
            .Connection = oCn
            .CommandText = sProcName
            .CommandType = CommandType.StoredProcedure
            .CommandTimeout = 0
        End With

        '//--- Add the parameters to the command
        AddParameters(oEnumerator, oCmd)

        ' Initialize the error code value
        ErrorCode = 0

        ' Open the connection
        oCn.Open()

        '//--- Execute the Stored Procedure
        oCmd.ExecuteNonQuery()

        '//--- Disconnect from the database
        Disconnect(oCn)

        '//--- Get the output parameter values if there were no errors
        If ErrorCode = 0 Then
            GetOutputParameterValues(oCmd)
        End If
    End Sub

    '//------------------------------------------------------------------------------------------------------------
    '// Public Method
    '// Overloaded:		Yes
    '// Return Value:	DataReader
    '// Purpose:		Executes a stored procedure.
    '//------------------------------------------------------------------------------------------------------------
    Public Function RunProcReader(ByVal sProcName As String) As SqlDataReader

        Dim dr As SqlDataReader                             '//--- Create a new sqlDataReader
        Dim oCmd As SqlCommand = New SqlCommand             '//--- Create a new SqlCommand
        Dim oCn As SqlConnection = Nothing                  '//--- Declare the SqlConnection
        Dim oSqlParameter As SqlParameter = Nothing         '//--- Declare a SqlParameter
        Dim oP As Parameter = Nothing                       '//--- Declare a Parameter
        '//--- Get an enumerator for the parameter array list
        Dim oEnumerator As IEnumerator = m_oParmList.GetEnumerator()

        '//--- Prepare connection to the database
        oCn = Connect()

        With oCmd
            '//--- Set the CommandText, ActiveConnection and the Command Type for the 
            '//--- SqlCommand Object
            .Connection = oCn
            .CommandText = sProcName
            .CommandType = CommandType.StoredProcedure
        End With

        '//--- Add the parameters to the command
        AddParameters(oEnumerator, oCmd)

        ' Open the connection
        oCn.Open()

        ' Execute the datareader, and close the connection
        dr = oCmd.ExecuteReader(CommandBehavior.CloseConnection)

        ' Get the output parameters
        GetOutputParameterValues(oCmd)

        '//--- Return the DataReader
        Return dr

    End Function

    '//---------------------------------------------------------------------------------------
    '// Private Method
    '// Return Value:	None
    '// Purpose:		Subroutine to add the parameters to an SQL stored procedure command
    '//---------------------------------------------------------------------------------------
    Private Sub AddParameters(ByVal oEnumerator As IEnumerator, ByVal oCmd As SqlCommand)
        Dim oP As Parameter
        Dim oSqlParameter As SqlParameter

        '//--- Loop through the Parameters in the ArrayList
        Do While (oEnumerator.MoveNext())
            oP = Nothing
            '//--- Get the current Parameter object
            oP = oEnumerator.Current
            '//--- Instantiate a SqlParameter object
            oSqlParameter = ConvertParameterToSqlParameter(oP)
            '//--- Add the SqlParameter object to the SqlCommand object
            oCmd.Parameters.Add(oSqlParameter)
        Loop
    End Sub

    '//---------------------------------------------------------------------------------------
    '// Private Method
    '// Return Value:	None
    '// Purpose:		Loop through and set the output parameters
    '//---------------------------------------------------------------------------------------
    Private Sub GetOutputParameterValues(ByVal oCmd As SqlCommand)
        '//--- declare the necessary working fields
        Dim checkParameter As SqlClient.SqlParameter
        Dim chkParameter As Parameter

        '//--- for each parameter in the command list
        For Each checkParameter In oCmd.Parameters
            '//--- if the parameter was an output parameter
            If checkParameter.Direction = ParameterDirection.Output Or checkParameter.Direction = ParameterDirection.InputOutput Then
                '//-- Find the parameter in the original parameter list
                For Each chkParameter In m_oParmList
                    '//--- If the parameter names are equal, assign the value of
                    '      the returned data to the original parameter value
                    If chkParameter.ParameterName = checkParameter.ParameterName Then
                        chkParameter.Value = checkParameter.Value
                    End If
                Next
            End If
        Next

    End Sub

    '//---------------------------------------------------------------------------------------
    '// Public Method
    '// Return Value:	Object
    '// Purpose:		return the value of an SQL server stored procedure outut parameter
    '//---------------------------------------------------------------------------------------
    Public Function ParameterValue(ByVal sParameterName As String) As Object
        Dim chkParameter As Parameter
        Dim returnValue As Object = Nothing

        '//-- Find the parameter in the original parameter list
        For Each chkParameter In m_oParmList
            '//--- If the parameter names are equal, assign the value of
            '      the returned data to the original parameter value
            If chkParameter.ParameterName = sParameterName Then
                returnValue = chkParameter.Value
            End If
        Next

        '//--- Return the appropriate parameters output value
        Return returnValue
    End Function

    '//---------------------------------------------------------------------------------------
    '// Public Method
    '// Overloaded:		Yes
    '// Return Value:	None
    '// Purpose:		Adds a parameter for a stored procedure.
    '//---------------------------------------------------------------------------------------
    Public Sub AddParameter(ByVal sParameterName As String, ByVal lSqlType As SqlDbType, ByVal iSize As Integer, ByVal sValue As Object, ByVal pdDirection As ParameterDirection)

        Dim oParam As Parameter = Nothing

        oParam = New Parameter(sParameterName, lSqlType, iSize, sValue, pdDirection)

        m_oParmList.Add(oParam)

    End Sub

    '//---------------------------------------------------------------------------------------
    '// Public Method
    '// Overloaded:		Yes
    '// Return Value:	None
    '// Purpose:		Adds a parameter for a stored procedure so that it can be referenced.
    '//                 after the call to get a changed output value.
    '//---------------------------------------------------------------------------------------
    Public Sub AddParameter(ByVal sParameter As Parameter)

        m_oParmList.Add(sParameter)

    End Sub

    '//---------------------------------------------------------------------------------------
    '// Private Method
    '// Return Value:	None
    '// Purpose:		Public Method that converts a Parameter to a SqlParameter
    '//---------------------------------------------------------------------------------------
    Private Function ConvertParameterToSqlParameter(ByVal oP As Parameter) As SqlParameter
        '//--- Instantiate a SqlParameter object
        Dim oSqlParameter As SqlParameter = New SqlParameter(oP.ParameterName, oP.DataType, oP.Size)

        With oSqlParameter
            .Value = oP.Value
            .Direction = oP.Direction
        End With

        Return oSqlParameter

    End Function

    '//---------------------------------------------------------------------------------------
    '// Private Method
    '// Return Value:	None
    '// Purpose:		Public Method that adds a parameter for a stored procedure.
    '//---------------------------------------------------------------------------------------
    Private Function Connect() As SqlConnection

        Dim oCn As SqlConnection = Nothing

        If (m_sConnectionString.Length > 0) Then
            oCn = New SqlConnection(m_sConnectionString)
        Else
            oCn = New SqlConnection("Server=" + m_sServer + ";User ID=" + m_sUsername + ";Password=" + m_sPassword + ";Database=" + m_sDatabase + ";")
        End If

        Return oCn

    End Function

    '//---------------------------------------------------------------------------------------
    '// Private Method
    '// Return Value:	None
    '// Purpose:		Close and destroy the Connection
    '//---------------------------------------------------------------------------------------
    Private Sub Disconnect(ByVal oCn As SqlConnection)
        '//--- Close the Connection if the Connection exists
        If (Not (oCn Is Nothing) And (oCn.State <> ConnectionState.Closed)) Then
            oCn.Close()
        End If

        '//--- Destroy the objects and release the memory
        oCn = Nothing
    End Sub

End Class

'//---------------------------------------------------------------------------------------
'// Public Class
'// Purpose:		The parameter class defines a parameter that will be passed 
'//				    to a stored procedure
'//---------------------------------------------------------------------------------------
Public Class Parameter

    Public DataType As SqlDbType            '//--- The datatype of the parameter
    Public Direction As ParameterDirection  '//--- The direction of the parameter
    Public ParameterName As String          '//--- The Name of the parameter
    Public Size As Integer                  '//--- The size of the parameter
    Public Value As Object                  '//--- The value of the parameter

    '//----------------------------------------------------------------------------------------------------
    '// Class Constructor (zero arguments)
    '// Overloaded:	No
    '//----------------------------------------------------------------------------------------------------
    Sub New(ByVal sParameterName As String, ByVal lDataType As SqlDbType, ByVal iSize As Integer, ByVal sValue As Object, ByVal pdDirection As ParameterDirection)
        ParameterName = sParameterName
        DataType = lDataType
        Size = iSize
        Value = sValue
        Direction = pdDirection
    End Sub

End Class

