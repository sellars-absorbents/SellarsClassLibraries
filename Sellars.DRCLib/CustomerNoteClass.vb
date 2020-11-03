Imports Sellars.SQL
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

Public Class CustomerNoteClass
    Inherits ClassBase

    Private _CustomerID As String = ""
    Private _Code As String = ""
    Private _Note As String = ""

    Public ReadOnly Property CustomerID() As String
        Get
            Return _CustomerID
        End Get
    End Property

    Public ReadOnly Property Code() As String
        Get
            Return _Code
        End Get
    End Property

    Public ReadOnly Property Note() As String
        Get
            Return _Note
        End Get
    End Property

    ' You have to initialize the class with a customer number 
    ' then you can add, delete, read and update notes for that customer
    Public Sub New(ByVal CustomerID As String)
        _CustomerID = CustomerID
    End Sub

    '*********************************************************************
    ' Add(NoteText)
    ' Adds a new note to the customer note table
    '*********************************************************************
    Public Sub Add(ByVal NoteText As String)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, _CustomerID, ParameterDirection.Input)
        oSQL.AddParameter("@NoteText", SqlDbType.NVarChar, 4000, NoteText, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddCustomerNote")
    End Sub

    '*********************************************************************
    ' Delete()
    ' Deletes the customer note for a customer.
    '*********************************************************************
    Public Sub Delete()
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, _CustomerID, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("DeleteCustomerNote")
    End Sub

    '*********************************************************************
    ' Read()
    ' Reads the customer note for a customer.
    '*********************************************************************
    Public Function Read() As String

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, _CustomerID, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Using dr As SqlDataReader = oSQL.RunProcReader("GetCustomerNote")
            ' Clear out the note field
            _Note = ""

            ' Assign the variables from the database to properties
            If dr.Read() Then
                _Note = dr("NoteText")
            Else
                _Note = ""
            End If
        End Using

        ' Return the note variable
        Return _Note
    End Function

    '*********************************************************************
    ' Update(NoteText)
    ' Updates a note text to the customer note table
    '*********************************************************************
    Public Sub Update(ByVal NoteText As String)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, _CustomerID, ParameterDirection.Input)
        oSQL.AddParameter("@NoteText", SqlDbType.NVarChar, 4000, NoteText, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateCustomerNote")
    End Sub

End Class
