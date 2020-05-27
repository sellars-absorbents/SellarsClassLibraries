'Imports Sellars.SQL
'Imports System.Collections.Generic
'Imports System.Configuration
'Imports System.Data
'Imports System.Data.SqlClient
'Imports System.Runtime.InteropServices
'Imports System.Text
'Imports System.Threading.Tasks

'Public Class StockAdjustment

'    ' MaxUpdate Sales Order Line Item functions
'    Declare Function ProcessTrans Lib "MAXTRAN2.DLL" (ByVal OH As Integer, ByRef Trndet As TRNDetItem) As Integer

'    ' MaxUpdate Sales Order Line Item Passing Structure
'    ' The PACK:=4 makes sure that numeric fields are set on 4 byte boundaries and is necessary to pass
'    ' the data correctly to the MaxUpdate function.
'    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=4)> _
'    Public Structure TRNDetItem
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=8)> Public ORDNUM() As Char   ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public LINNUM() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public DELNUM() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public STATUS() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=20)> Public CUSTID() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public PRTNUM() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=6)> Public EDILIN() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public TAXABL() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=32)> Public GLXREF() As Char ' String
'        Public CURDUE As Integer ' Btrieve Date
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public QTLINE() As Char ' String
'        Public ORGDUE As Integer ' Btrieve Date
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public QTDEL() As Char ' String
'        Public CUSDUE As Integer ' Btrieve Date
'        Public PROBAB As Integer
'        Public SHPDTE As Integer ' Btrieve Date
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public FILL04() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public SLSUOM() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=25)> Public REFRNC() As Char ' String
'        Public PRICE As Double ' IEEE Float
'        Public ORGQTY As Double ' IEEE Float
'        Public CURQTY As Double ' IEEE Float
'        Public BCKQTY As Double ' IEEE Float
'        Public SHPQTY As Double ' IEEE Float
'        Public CURSHP As Double ' IEEE Float
'        Public DUEQTY As Double ' IEEE Float
'        Public INVQTY As Double ' IEEE Float
'        Public DISC As Single ' IEEE Float
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public STYPE() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public PRNT() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public AKPRNT() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=8)> Public STK() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=3)> Public COCFLG() As Char ' String
'        Public FORCUR As Double ' IEEE Float
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public HSTAT() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public SLSREP() As Char ' String
'        Public COMMIS As Single ' IEEE Float
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=10)> Public DRPSHP() As Char ' String
'        Public QUMQTY As Single ' IEEE Float
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public TAXCDE1() As Char ' String
'        Public TAX1 As Double ' IEEE Float
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public TAXCDE2() As Char ' String
'        Public TAX2 As Double ' IEEE Float
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public TAXCDE3() As Char ' String
'        Public TAX3 As Double ' IEEE Float
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=3)> Public MCOMP() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=3)> Public MSITE() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=15)> Public UDFKEY() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=25)> Public UDFREF() As Char ' String
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public DEXPFLG() As Char ' String (adChar)
'        Public COST As Double
'        Public MARKUP As Double
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=8)> Public QTORD() As Char 'String
'        Public XDFINT As Integer
'        Public XDFFLT As Double
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public XDFBOL() As Char
'        Public XDFDTE As Integer
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=100)> Public XDFTXT() As Char
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=50)> Public FILLER() As Char
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=100)> Public CREATEDBY() As Char
'        Public CREATIONDATE As Integer
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=100)> Public MODIFIEDBY() As Char
'        Public MODIFICATIONDATE As Integer
'        Public BOKDTE As Integer
'        Public DBKDTE As Integer
'        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=3)> Public REVLEV() As Char
'    End Structure

'End Class
