Option Strict Off
Option Explicit On 

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Math
Imports System.Windows.Forms

Namespace Windows
    Public Class QCDataGridTextBoxColumn
        Inherits DataGridTextBoxColumn

        ' Note, these values need to synch up with the values in the FieldTypeEnum defined in the
        ' QCParameters class through all the items that it has in it.
        Public Enum FieldTypeEnum
            Unassigned = 0
            Weight = 1
            OpSideBulk = 2
            DriveSideBulk = 3
            MDT = 4
            MDTE = 5
            CDTD = 6
            CDTW = 7
            CDTC = 8
            TT = 9
            ZPeel = 10
            Contamination = 11
            TWA = 12
            C_Percentage = 13
            WD = 14
            BPDate = 15
            BPReelNo = 16
        End Enum

        Private m_OwnerObject As Object
        Private m_FieldType As Integer = 0
        Private Const FieldTypeDescription As String = "Type of data being displayed."

        'Properties
        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(FieldTypeDescription)> _
        Public Property FieldType() As FieldTypeEnum
            Get
                Return m_FieldType
            End Get
            Set(ByVal Value As FieldTypeEnum)
                m_FieldType = Value
            End Set
        End Property

        'Constructors
        'Events
        'Methods
        Public Sub New(ByVal passOwnerObject As Object)
            'Warning: Implementation not found
            m_OwnerObject = passOwnerObject
        End Sub
        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)

            ' the idea is to conditionally set the foreBrush and/or backbrush
            ' depending upon some criteria on the cell value
            Try
                Dim o As Object
                o = Me.GetColumnValueAtRow(source, rowNum)
                If (Not (o) Is Nothing) Then
                    Dim c As Decimal
                    c = o
                    If c = 0 Then
                        backBrush = GrayBrush(bounds)
                    Else
                        backBrush = GetBackgroundBrush(c, bounds)
                    End If
                    foreBrush = New SolidBrush(Color.Black)
                End If
            Catch ex As Exception
                ' empty catch
            Finally
                ' make sure the base class gets called to do the drawing with
                ' the possibly changed brushes
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try

        End Sub

        Private Function GetBackgroundBrush(ByVal passValue As Decimal, ByVal bounds As Rectangle) As Brush
            Dim backBrush As Brush = Nothing

            ' If the field type is percent c compare to three decimal places 
            If FieldType = FieldTypeEnum.C_Percentage Then
                Select Case Round(passValue, 3)
                    Case Is < Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 3)
                        backBrush = RedBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 3) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 3) - 0.001
                        backBrush = YellowBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 3) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 3) - 0.001
                        backBrush = GreenBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 3) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 3) - 0.001
                        backBrush = YellowBrush(bounds)
                    Case Is >= Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 3)
                        backBrush = RedBrush(bounds)
                End Select
            ElseIf FieldType = FieldTypeEnum.WD Then
                ' If the field type is wd compare to two decimal places 
                Select Case Round(passValue, 2)
                    Case Is < Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 2)
                        backBrush = RedBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 2) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 2) - 0.01
                        backBrush = YellowBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 2) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 2) - 0.01
                        backBrush = GreenBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 2) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 2) - 0.01
                        backBrush = YellowBrush(bounds)
                    Case Is >= Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 2)
                        backBrush = RedBrush(bounds)
                End Select
            ElseIf FieldType = FieldTypeEnum.Weight Or FieldType = FieldTypeEnum.OpSideBulk Or FieldTypeEnum.DriveSideBulk Then
                ' If the field type is weight or bulk compare to one decimal place, 
                ' otherwise compare to whole numbers
                Select Case Round(passValue, 1)
                    Case Is < Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 1)
                        backBrush = RedBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 1) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 1) - 0.1
                        backBrush = YellowBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 1) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 1) - 0.1
                        backBrush = GreenBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 1) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 1) - 0.1
                        backBrush = YellowBrush(bounds)
                    Case Is >= Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 1)
                        backBrush = RedBrush(bounds)
                End Select
            Else
                Select Case Round(passValue, 0)
                    Case Is < Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 0)
                        backBrush = RedBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, 0) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 0) - 1
                        backBrush = YellowBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, 0) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 0) - 1
                        backBrush = GreenBrush(bounds)
                    Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, 0) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 0) - 1
                        backBrush = YellowBrush(bounds)
                    Case Is >= Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, 0)
                        backBrush = RedBrush(bounds)
                End Select
            End If

            Return backBrush
        End Function

    End Class
End Namespace

