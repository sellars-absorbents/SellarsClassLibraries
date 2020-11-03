Imports System.Data
Imports System.Data.SqlClient

Public Class ControlValues

    Private Shared lockObject As New Object()

    Private err As Boolean = False
    Private m_errorMessage As String = ""
    Private m_text As String = ""

    Public Property [Error]() As Boolean
        Get
            Return err
        End Get
        Private Set(ByVal value As Boolean)
            err = value
        End Set
    End Property

    Public Property ErrorMessage() As String
        Get
            Return m_errorMessage
        End Get
        Private Set(ByVal value As String)
            m_errorMessage = value
        End Set
    End Property

    Public Property Text() As String
        Get
            Return m_text
        End Get
        Private Set(ByVal value As String)
            m_text = value
        End Set
    End Property

    Public Sub New()
    End Sub

    Public Function Increment(ByVal conn As SqlConnection, ByVal Key As String) As String
        Dim val As String = ""

        ' Need to lock the object across all incrementing threads to make sure that 
        ' we never have thread objects using the same incremented key value
        SyncLock (lockObject)
            ' Read the designated control value
            val = Read(conn, Key)

            ' increment the value by 1
            val = CStr(Convert.ToInt32(val) + 1)

            ' Post the new value
            WriteValue(conn, Key, val)
        End SyncLock

        Return val
    End Function

    Public Function Increment(ByVal conn As SqlConnection, ByRef tran As SqlTransaction, ByVal Key As String) As String
        Dim val As String = ""

        ' Need to lock the object across all incrementing threads to make sure that 
        ' we never have thread objects using the same incremented key value
        SyncLock (lockObject)
            ' Read the designated control value
            val = Read(conn, tran, Key)

            ' increment the value by 1
            val = CStr(Convert.ToInt32(val) + 1)

            ' Post the new value
            WriteValue(conn, tran, Key, val)
        End SyncLock

        Return val
    End Function

    Public Function Read(ByVal conn As SqlConnection, ByRef tran As SqlTransaction, ByVal Key As String) As String
        Dim command As String = "SELECT @Value = ISNULL(Value, '')" & vbTab & "from dbo.ControlValues where [Key] = @Key"

        Using cmd As New SqlCommand(command, conn)
            cmd.CommandType = CommandType.Text
            cmd.Transaction = tran
            cmd.Parameters.Add(New SqlParameter("@Key", Key))

            Dim parmValue As New SqlParameter("@Value", SqlDbType.NVarChar, -1)
            parmValue.Direction = ParameterDirection.Output
            parmValue.Value = Nothing
            cmd.Parameters.Add(parmValue)

            Try
                cmd.ExecuteNonQuery()
                Text = parmValue.Value.ToString()
            Catch ex As Exception
                [Error] = True
                ErrorMessage = ex.Message
                Text = ""
            End Try
        End Using

        Return Text
    End Function

    Public Function Read(ByVal conn As SqlConnection, ByVal Key As String) As String
        Dim rtnValue As String = ""
        Dim command As String = "SELECT @Value = ISNULL(Value, '')" & vbTab & "from dbo.ControlValues where [Key] = @Key"

        Using cmd As New SqlCommand(command, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@Key", Key))

            Dim parmValue As New SqlParameter("@Value", SqlDbType.NVarChar, -1)
            parmValue.Direction = ParameterDirection.Output
            parmValue.Value = Nothing
            cmd.Parameters.Add(parmValue)

            Try
                cmd.ExecuteNonQuery()
                rtnValue = parmValue.Value.ToString()
            Catch ex As Exception
                [Error] = True
                ErrorMessage = ex.Message
                rtnValue = ""
            End Try
        End Using

        Return rtnValue
    End Function

    Public Function ReadByte(ByVal conn As SqlConnection, ByVal Key As String) As String
        Dim rtnValue As String = Nothing
        Dim command As String = "SELECT @Value = isnull(FileData, '') from dbo.ControlValues where [Key] = @Key"

        Using cmd As New SqlCommand(command, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@Key", Key))

            Dim parmValue As New SqlParameter("@Value", SqlDbType.NVarChar, -1)
            parmValue.Direction = ParameterDirection.Output
            parmValue.Value = Nothing
            cmd.Parameters.Add(parmValue)

            Try
                cmd.ExecuteNonQuery()
                rtnValue = parmValue.Value.ToString()
            Catch ex As Exception
                [Error] = True
                ErrorMessage = ex.Message
                rtnValue = ""
            End Try
        End Using

        Return rtnValue
    End Function

    Public Function WriteValue(ByVal conn As SqlConnection, ByRef tran As SqlTransaction, ByVal Key As String, ByVal Value As String) As String
        Dim rtnValue As String = ""
        Dim command As String = "update dbo.ControlValues Set Value = @Value where [Key] = @Key"

        Using cmd As New SqlCommand(command, conn)
            cmd.Transaction = tran
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@Key", Key))
            cmd.Parameters.Add(New SqlParameter("@Value", Value))

            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                [Error] = True
                ErrorMessage = ex.Message
                rtnValue = ""
            End Try
        End Using

        Return rtnValue
    End Function

    Public Function WriteValue(ByVal conn As SqlConnection, ByVal Key As String, ByVal Value As String) As String
        Dim rtnValue As String = ""
        Dim command As String = "update dbo.ControlValues Set Value = @Value where [Key] = @Key"

        Using cmd As New SqlCommand(command, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@Key", Key))
            cmd.Parameters.Add(New SqlParameter("@Value", Value))

            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                [Error] = True
                ErrorMessage = ex.Message
                rtnValue = ""
            End Try
        End Using

        Return rtnValue
    End Function

    Public Function Write(ByVal conn As SqlConnection, ByVal Key As String, ByVal Value As String) As String
        Dim rtnValue As String = ""
        Dim command As String = "update dbo.ControlValues Set FileData = @Value where [Key] = @Key"

        Using cmd As New SqlCommand(command, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@Key", Key))
            cmd.Parameters.Add(New SqlParameter("@Value", Value))

            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                [Error] = True
                ErrorMessage = ex.Message
                rtnValue = ""
            End Try
        End Using

        Return rtnValue
    End Function

End Class
