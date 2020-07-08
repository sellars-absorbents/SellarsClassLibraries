Imports System.Xml.Linq

Public Class MAXTransaction
    Inherits ClassBase

#Region "Properties"
    Public Property TYPE_39() As Char = " "c
    Public Property SUBTYPE_39() As Char = " "c 'blank space (not empty string) required for specific types of transactions - MAX Update User Guide
    Public Property RCVSTK_39() As String = ""
    Public Property ISSSTK_39() As String = ""
    Public Property PRTNUM_39() As String = ""
    Public Property ORDNUM_39() As String = ""
    Public Property LINNUM_39() As String = ""
    Public Property DELNUM_39() As String = ""
    Public Property TNXQTY_39() As Decimal = 0D
    Public Property TNXDTE_39() As DateTime = New DateTime(1, 1, 1) ' null date in MAX is 0001-01-01 ' MMddyy required format - MAX Update User Guide
    Public Property TNXTME_39() As DateTime = New DateTime(1, 1, 1) ' hhmmss required format - MAX Update User Guide
    Public Property REFDSC_39() As String = ""
    Public Property GLREF_39() As String = ""
    Public Property OPRSEQ_39() As String = ""
    Public Property NXTOPR_39() As String = ""
    Public Property RUNACT_39() As Decimal = 0D
    Public Property SETACT_39() As Decimal = 0D
    Public Property SHIFT_39() As Integer = 0
    Public Property LOCATOR_39() As String = ""
    Public Property EMPID_39() As String = ""
    Public Property TICKET_39() As String = ""
    Public Property STARTTIME_39() As Date = New Date(1, 1, 1)
    Public Property ENDTIME_39() As Date = New Date(1, 1, 1)
    Public Property SETUPTIME_39() As Decimal = 0D
    Public Property ELAPSED_39() As Decimal = 0D
    Public Property ASCRAP_39() As Decimal = 0D
    Public Property REASON_39() As Char = " "c
    Public Property USERNAME_39() As String = ""
    Public Property UDFKEY_39() As String = ""
    Public Property UDFREF_39() As String = ""
    Public Property ASSCODE_39() As Char = " "c
    Public Property LOT_39() As String = ""
    Public Property SERIAL_39() As String = ""
    Public Property TERMINAL_39() As String = ""
    Public Property QCODE_39() As Char = " "c
    Public Property EXPDATE_39() As Date = New Date(1, 1, 1)
    Public Property DEFECT_39() As String = ""
    Public Property RECPL_39() As Char = " "c
    Public Property DISPOSITION_39() As String = ""
    Public Property CLASS_39() As String = ""
#End Region

#Region "Methods"
    Public Sub GenericTransaction()
        Dim transaction As XDocument = CreateXDocumentForTransaction()
        Dim errorMessage As String = ""
        Dim retValue As Integer = ProcessTransXML(transaction.ToString(), errorMessage)

        If retValue = 0 Then
            Throw New ApplicationException("There was an error processing the transaction for order number " + ORDNUM_39 + ". Error: " + errorMessage)
        End If
    End Sub

    Public Sub Ship()
        If TYPE_39 <> "S" Then
            Throw New ApplicationException("The Type must be S for a ship transaction")
        End If

        Dim transaction As XDocument = CreateXDocumentForTransaction()
        Dim errorMessage As String = ""
        Dim retValue As Integer = ProcessTransXML(transaction.ToString(), errorMessage)

        If retValue = 0 Then
            Throw New ApplicationException("There was an error processing the ship transaction for order number " + ORDNUM_39 + ". Error: " + errorMessage)
        End If
    End Sub

    Public Sub StockTransfer()
        If TYPE_39 <> "F" OrElse SUBTYPE_39 <> " " Then
            Throw New ApplicationException("The Type must be S and the subtype must be one blank space for a stock transfer transaction")
        End If

        Dim transaction As XDocument = CreateXDocumentForTransaction()
        Dim errorMessage As String = ""
        Dim retValue As Integer = ProcessTransXML(transaction.ToString(), errorMessage)

        If retValue = 0 Then
            Throw New ApplicationException("There was an error processing the stock transfer transaction for part " + PRTNUM_39 + ". Error: " + errorMessage)
        End If
    End Sub

    Public Sub Receipt()
        If TYPE_39 <> "R" OrElse SUBTYPE_39 <> "P" Then
            Throw New ApplicationException("The Type must be R and the subtype must be P for a receipt transaction")
        End If

        Dim transaction As XDocument = CreateXDocumentForTransaction()
        Dim errorMessage As String = ""
        Dim retValue As Integer = ProcessTransXML(transaction.ToString(), errorMessage)

        If retValue = 0 Then
            Throw New ApplicationException("There was an error processing the receipt transaction for order number " + ORDNUM_39 + ". Error: " + errorMessage)
        End If
    End Sub

    Public Sub Adjustment()
        If TYPE_39 <> "A" OrElse SUBTYPE_39 <> " " Then
            Throw New ApplicationException("The Type must be A and the subtype must be one blank space for an adjustment transaction")
        End If

        Dim transaction As XDocument = CreateXDocumentForTransaction()
        Dim errorMessage As String = ""
        Dim retValue As Integer = ProcessTransXML(transaction.ToString(), errorMessage)

        If retValue = 0 Then
            Throw New ApplicationException("There was an error processing the adjustment transaction for part " + PRTNUM_39 + ". Error: " + errorMessage)
        End If
    End Sub

    Private Function CreateXDocumentForTransaction() As XDocument
        Dim transactionXML As XDocument =
            <?xml version="1.0" encoding="utf-8"?>
            <eMAXExact>
                <MAX_Transaction_Table>
                    <MAX_Transaction>
                        <TYPE_39><%= TYPE_39() %></TYPE_39>
                        <SUBTYPE_39><%= SUBTYPE_39() %></SUBTYPE_39>
                        <RCVSTK_39><%= RCVSTK_39().GetFixedLengthString(8) %></RCVSTK_39>
                        <ISSSTK_39><%= ISSSTK_39().GetFixedLengthString(8) %></ISSSTK_39>
                        <PRTNUM_39><%= PRTNUM_39().GetFixedLengthString(30) %></PRTNUM_39>
                        <ORDNUM_39><%= ORDNUM_39().GetFixedLengthString(8) %></ORDNUM_39>
                        <LINNUM_39><%= LINNUM_39().GetFixedLengthString(2) %></LINNUM_39>
                        <DELNUM_39><%= DELNUM_39().GetFixedLengthString(2) %></DELNUM_39>
                        <TNXQTY_39><%= TNXQTY_39() %></TNXQTY_39>
                        <TNXDTE_39><%= TNXDTE_39().ToString("MMddyy") %></TNXDTE_39>
                        <TNXTME_39><%= TNXTME_39().ToString("hhmmss") %></TNXTME_39>
                        <REFDSC_39><%= REFDSC_39().GetFixedLengthString(20) %></REFDSC_39>
                        <GLREF_39><%= GLREF_39().GetFixedLengthString(3) %></GLREF_39>
                        <OPRSEQ_39><%= OPRSEQ_39().GetFixedLengthString(4) %></OPRSEQ_39>
                        <NXTOPR_39><%= NXTOPR_39().GetFixedLengthString(4) %></NXTOPR_39>
                        <RUNACT_39><%= RUNACT_39() %></RUNACT_39>
                        <SETACT_39><%= SETACT_39() %></SETACT_39>
                        <SHIFT_39><%= SHIFT_39() %></SHIFT_39>
                        <LOCATOR_39><%= LOCATOR_39().GetFixedLengthString(10) %></LOCATOR_39>
                        <EMPID_39><%= EMPID_39().GetFixedLengthString(7) %></EMPID_39>
                        <TICKET_39><%= TICKET_39().GetFixedLengthString(6) %></TICKET_39>
                        <STARTTIME_39><%= MakeDate(STARTTIME_39()) %></STARTTIME_39>
                        <ENDTIME_39><%= MakeDate(ENDTIME_39()) %></ENDTIME_39>
                        <SETUPTIME_39><%= SETUPTIME_39() %></SETUPTIME_39>
                        <ELAPSED_39><%= ELAPSED_39() %></ELAPSED_39>
                        <ASCRAP_39><%= ASCRAP_39() %></ASCRAP_39>
                        <REASON_39><%= REASON_39() %></REASON_39>
                        <USERNAME_39><%= USERNAME_39().GetFixedLengthString(100) %></USERNAME_39>
                        <UDFKEY_39><%= UDFKEY_39().GetFixedLengthString(15) %></UDFKEY_39>
                        <UDFREF_39><%= UDFREF_39().GetFixedLengthString(25) %></UDFREF_39>
                        <ASSCODE_39><%= ASSCODE_39() %></ASSCODE_39>
                        <LOT_39><%= LOT_39().GetFixedLengthString(30) %></LOT_39>
                        <SERIAL_39><%= SERIAL_39().GetFixedLengthString(30) %></SERIAL_39>
                        <TERMINAL_39><%= TERMINAL_39().GetFixedLengthString(15) %></TERMINAL_39>
                        <QCODE_39><%= QCODE_39() %></QCODE_39>
                        <EXPDATE_39><%= MakeDate(EXPDATE_39()) %></EXPDATE_39>
                        <DEFECT_39><%= DEFECT_39().GetFixedLengthString(4) %></DEFECT_39>
                        <RECPL_39><%= RECPL_39() %></RECPL_39>
                        <DISPOSITION_39><%= DISPOSITION_39().GetFixedLengthString(15) %></DISPOSITION_39>
                        <CLASS_39><%= CLASS_39().GetFixedLengthString(159) %></CLASS_39>
                    </MAX_Transaction>
                </MAX_Transaction_Table>
            </eMAXExact>

        Return transactionXML
    End Function

    ' This function is the initialization routine that needs to be called before any MaxUpdateXML function is called
    Public Overridable Sub InitializeMax(ByVal connStr As String, ByVal comName As String, ByVal licPath As String, ByVal logPath As String, ByVal log As Boolean)
        ClassBase.Initialize(connStr, comName, licPath, logPath, log)
    End Sub

    Protected Overridable Function ProcessTransXML(ByVal xml As String, ByVal showErrorMessages As Boolean) As Integer
        Return ClassBase.MAXUpdate.ProcessTransXML(xml, showErrorMessages)
    End Function
#End Region

End Class
