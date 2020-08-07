Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Xml
Imports System.Web.Services
Imports System.Web.Services.Protocols

Public Class ZonePriceList
    Private rs As ReportService.ReportingService2010 = New ReportService.ReportingService2010()
    Private rsExec As ReportExecutionService.ReportExecutionService = New ReportExecutionService.ReportExecutionService()

    Private _ReportURL As String = ""
    Private m_renderedReport As Byte()()
    Private m_delegate As Graphics.EnumerateMetafileProc = Nothing
    Private m_currentPageStream As MemoryStream
    Private m_metafile As Metafile = Nothing
    Private m_numberOfPages As Integer = 0
    Private m_currentPrintingPage As Integer = 0
    Private m_lastPrintingPage As Integer = 0

    Private Property RenderedReport As Byte()()
        Get
            Return m_renderedReport
        End Get
        Set(value As Byte()())
            m_renderedReport = value
        End Set
    End Property

    Public Sub New(ByVal ReportServer As String, ByVal ReportURL As String, ByVal Username As String, ByVal Password As String, ByVal Domain As String)
        rs.Credentials = New NetworkCredential(Username, Password, Domain)
        rsExec.Credentials = New NetworkCredential(Username, Password, Domain)
        rsExec.Url = ReportServer
        _ReportURL = ReportURL
    End Sub

    Public Function PrintToPDF(ByVal State As String, ByVal Format As String) As MemoryStream
        Dim rtnReport As MemoryStream = Nothing
        Dim historyID As String = Nothing
        Dim deviceInfo As String = Nothing
        Dim results As Byte() = Nothing
        Dim encoding As String = String.Empty
        Dim mimeType As String = String.Empty
        Dim extension As String = String.Empty
        Dim warnings As ReportExecutionService.Warning() = Nothing
        Dim streamIDs As String() = Nothing

        ' Define variables needed for the GetParameters() method
        Dim _historyID As String = Nothing
        Dim _forRendering As Boolean = False

        Dim _values As ReportService.ParameterValue() = Nothing
        Dim _credentials As ReportService.DataSourceCredentials() = Nothing
        Dim _parameters As ReportService.ParameterValue() = Nothing

        ' Load the selected report
        Dim ei As ReportExecutionService.ExecutionInfo = rsExec.LoadReport(_ReportURL, historyID)

        ' prepare the report parameters
        Dim parameters As ReportExecutionService.ParameterValue() = SetParameters(State)

        ' Prepare device info to render PDF at 300 DPI
        Dim sb As New System.Text.StringBuilder(1024)
        Dim xr As System.Xml.XmlWriter = XmlWriter.Create(sb)
        xr.WriteStartElement("DeviceInfo")
        xr.WriteElementString("DpiX", "300")
        xr.WriteElementString("DpiY", "300")
        xr.Close()

        deviceInfo = sb.ToString()

        rsExec.SetExecutionParameters(parameters, "en-us")
        results = rsExec.Render(Format, deviceInfo, extension, mimeType, encoding, warnings, streamIDs)

        ' put the results into the return memorystream
        rtnReport = New MemoryStream(results)

        ' return the report memorystream object
        Return rtnReport
    End Function

    Private Function SetParameters(ByVal State As String) As ReportExecutionService.ParameterValue()
        Dim parameters(3) As ReportExecutionService.ParameterValue

        ' Set the warehouse parameter
        parameters(0) = New ReportExecutionService.ParameterValue()
        parameters(0).Label = "State"
        parameters(0).Name = "State"
        parameters(0).Value = State.ToUpper

        parameters(1) = New ReportExecutionService.ParameterValue()
        parameters(1).Label = "TargetDate"
        parameters(1).Name = "TargetDate"
        parameters(1).Value = Date.Now()

        ' return the parameters to the calling module
        Return parameters
    End Function
End Class
