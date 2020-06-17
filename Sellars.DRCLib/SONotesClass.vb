Imports System.Data
Imports System.Data.SqlClient

Public Class SONotesClass
    Inherits ClassBase

    Private _NoteCount As Integer = 0
    Private _StartPos As Integer = 0
    Private _CurrPos As Integer = 0

    Public Function Read(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal PrintType As String) As String
        ' Define a stringbuilder variable to hold all the notes returned from
        ' the sales order note master table
        Dim strNote As New System.Text.StringBuilder

        ' Get all the notes that match the criteria
        OpenMaxConnection()

        Dim strSQL As String = "Select COMNT_30 " &
                               "From ""SO_Note"" " &
                               "Where ORDNUM_30 = '" & ORDNUM & "' " &
                               "and LINNUM_30 = '" & LINNUM & "' " &
                               "and DELNUM_30 = '" & DELNUM & "' " &
                               "and CODE_30 = '" & PrintType & "' " &
                               "order by COMNUM_30"

        Using cmd As New SqlCommand(strSQL, MaxConnection)
            Using myReader As SqlDataReader = cmd.ExecuteReader()
                While myReader.Read()
                    If Trim(myReader("COMNT_30")) = "" Then
                        strNote.Append(vbCrLf & vbCrLf)
                    Else
                        strNote.Append(Trim(myReader("COMNT_30")))
                    End If
                End While
            End Using
        End Using

        CloseMaxConnection()

        ' Return the note string of all the notes read in
        Return strNote.ToString
    End Function

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal Both As String, ByVal InvoiceOnly As String, ByVal OrderOnly As String, ByVal Neither As String)
        ' OPen the max connection
        OpenMaxConnection()

        ' Delete all existing notes from the database
        Dim strSQL As String = "Delete From ""SO_Note"" " &
                               "Where ORDNUM_30 = '" & ORDNUM & "' " &
                               "and LINNUM_30 = '" & LINNUM & "' " &
                               "and DELNUM_30 = '" & DELNUM & "'"

        Using cmd As New SqlCommand(strSQL, MaxConnection)
            cmd.ExecuteNonQuery()
        End Using

        ' Initialize the _NoteCount variable to zero
        _NoteCount = 0

        ' Call the routine to add the Print on Both Invoice and Orders notes
        Add(ORDNUM, LINNUM, DELNUM, Both, "B")

        ' Call the routine to add the Print on Invoice Only notes
        Add(ORDNUM, LINNUM, DELNUM, InvoiceOnly, "I")

        ' Call the routine to add the Print on Order Only notes
        Add(ORDNUM, LINNUM, DELNUM, OrderOnly, "O")

        ' Call the routine to add the Print on Neither notes
        Add(ORDNUM, LINNUM, DELNUM, Neither, " ")

        ' Close the max database connection
        CloseMaxConnection()
    End Sub

    Private Sub Add(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal Notes As String, ByVal Type As String)
        ' If there are no notes to be added for that type, exit the routine
        If Notes.Trim = "" Then
            Exit Sub
        End If

        ' Split the record into 50 byte records, OR whenever a carriage return is encountered
        ' if a blank line entered, make sure that goes through as well
        _StartPos = 0
        While _StartPos < Notes.Trim.Length
            AddNote(ORDNUM, LINNUM, DELNUM, GetNoteLine(Notes), Type)
        End While
    End Sub

    Private Function GetNoteLine(ByVal Notes As String) As String
        ' Initialize necessary local variables
        Dim bStop As Boolean = False
        Dim charCount As Integer = 0
        Dim retString As New System.Text.StringBuilder

        ' If we encounter a carriage return, end the line and skip the carriage return
        ' Otherwise accumulate up to 50 characters, and return them
        While bStop = False And charCount < 50 And _StartPos < Notes.Trim.Length
            If Notes.Substring(_StartPos, 1) = vbCr Then
                bStop = True
                ' Increment by 2 to avoid the line feed as well
                _StartPos += 2
            Else
                retString.Append(Notes.Substring(_StartPos, 1))
                _StartPos += 1
                charCount += 1
            End If
        End While

        ' Return the information designated as a line of data to write to the 
        ' SONotes table
        Return retString.ToString
    End Function

    Private Sub AddNote(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal Note As String, ByVal Type As String)
        ' Increment the note count variable
        _NoteCount += 1

        ' Insert the new note into the database
        Dim strSQL As String = "insert into ""SO_Note"" (ORDNUM_30, LINNUM_30, DELNUM_30, COMNUM_30, CODE_30, COMNT_30, CUSTID_30, PIDCOD_30, MCOMP_30, MSITE_30, UDFKEY_30, UDFREF_30, XDFDTE_30, RECTYP_30) values (" &
                               "'" & ORDNUM & "','" & LINNUM & "','" & DELNUM & "','" & _NoteCount.ToString.PadLeft(2, "0") & "','" & Type & "','" & Note & "', '', '', '', '', '', '', GETDATE(), 'ST') "

        Using cmd As New SqlCommand(strSQL, MaxConnection)
            cmd.ExecuteNonQuery()
        End Using
    End Sub
End Class
