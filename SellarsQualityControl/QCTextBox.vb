Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Math
Imports System.Windows.Forms

Namespace Windows
    Public Class QCTextBox
        Inherits System.Windows.Forms.TextBox

#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call
            Me.SetStyle(ControlStyles.UserPaint, True)
        End Sub

        Public Sub New(ByVal passOwnerObject As Object)
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            m_OwnerObject = passOwnerObject

            'Add any initialization after the InitializeComponent() call
            Me.SetStyle(ControlStyles.UserPaint, True)
        End Sub

        'UserControl1 overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container
        End Sub

#End Region

        Private m_AllowNullValue As Boolean
        Private m_AllowZeroValue As Boolean
        Private m_SetBackgroundColor As Boolean
        Private m_HasFocus As Boolean
        Private m_FieldType As Integer = 0
        Private m_OwnerObject As Object
        Private m_CheckGreaterThanObject As Object
        Private m_CheckLessThanObject As Object
        Private m_CheckGreaterThan As Boolean = False
        Private m_CheckLessThan As Boolean = False
        Private m_ErrorMessage As String

        Private Const AllowNullValueDescription As String = "Allow null values as valid input."
        Private Const AllowZeroValueDescription As String = "Allow zero values as valid input."
        Private Const SetBackgroundColorDescription As String = "Set Background color based on value."
        Private Const OwnerDescription As String = "Parent object."
        Private Const FieldTypeDescription As String = "Type of data being displayed."
        Private Const CheckGreaterThanDescription As String = "Whether or not to compare the value to the greater than object's value."
        Private Const CheckLessThanDescription As String = "Whether or not to compare the value to the less than object's value."
        Private Const CheckGreaterThanObjectDescription As String = "Object whose value must be greater than the value contained in this item."
        Private Const CheckLessThanObjectDescription As String = "Object whose value must be less than the value contained in this item."

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
            Pcntc = 13
            WD = 14
            BPDate = 15
            BPReelNo = 16
        End Enum

        ' This is set up to show how many decimal digits to accept after the decimal point
        ' it must be kept in synch with the FieldTypeEnum with the same number of items
        ' This should also match how each field is defined in the SQL database with the decimal(18,x) where x equals
        ' the value for each of the digits in the declaration for each item below.
        Private DecimalDigits() As Short = {0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 3, 0, 0}

        'dim FieldNames() as String = ("Unassigned", "Weight", "Bulk", "MDT", "MDTE", "CDTD", "CDTC", "TT", "ZPeel", "BPDate", "BPReelNo")
        
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

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(OwnerDescription)> _
        Public Property Owner() As Object
            Get
                Return m_OwnerObject
            End Get
            Set(ByVal Value As Object)
                m_OwnerObject = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(AllowNullValueDescription)> _
        Public Property AllowNullValue() As Boolean
            Get
                Return m_AllowNullValue
            End Get
            Set(ByVal Value As Boolean)
                m_AllowNullValue = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(AllowNullValueDescription)> _
        Public Property HasFocus() As Boolean
            Get
                Return m_HasFocus
            End Get
            Set(ByVal Value As Boolean)
                m_HasFocus = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(AllowZeroValueDescription)> _
        Public Property AllowZeroValue() As Boolean
            Get
                Return m_AllowZeroValue
            End Get
            Set(ByVal Value As Boolean)
                m_AllowZeroValue = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(CheckGreaterThanDescription)> _
        Public Property CheckGreaterThan() As Boolean
            Get
                Return m_CheckGreaterThan
            End Get
            Set(ByVal Value As Boolean)
                m_CheckGreaterThan = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(CheckLessThanDescription)> _
        Public Property CheckLessThan() As Boolean
            Get
                Return m_CheckLessThan
            End Get
            Set(ByVal Value As Boolean)
                m_CheckLessThan = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(CheckGreaterThanObjectDescription)> _
        Public Property CheckGreaterThanObject() As Object
            Get
                Return m_CheckGreaterThanObject
            End Get
            Set(ByVal Value As Object)
                m_CheckGreaterThanObject = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(CheckLessThanObjectDescription)> _
        Public Property CheckLessThanObject() As Object
            Get
                Return m_CheckLessThanObject
            End Get
            Set(ByVal Value As Object)
                m_CheckLessThanObject = Value
            End Set
        End Property

        <System.ComponentModel.Category("Behavior"), _
         System.ComponentModel.Description(SetBackgroundColorDescription)> _
        Public Property SetBackgroundColor() As Boolean
            Get
                Return m_SetBackgroundColor
            End Get
            Set(ByVal Value As Boolean)
                m_SetBackgroundColor = Value
            End Set
        End Property

        <System.ComponentModel.Browsable(False)> _
        Public ReadOnly Property Valid() As Boolean
            Get
                Return IsValid(Text)
            End Get
        End Property

        <System.ComponentModel.Browsable(False)> _
                Public ReadOnly Property FieldEnumName() As String
            Get
                Dim tmpFieldTypeEnum As FieldTypeEnum
                Dim intIndex As Integer = FieldType

                tmpFieldTypeEnum = CType(intIndex, FieldTypeEnum)
                Return tmpFieldTypeEnum.ToString()
            End Get
        End Property

        <System.ComponentModel.Browsable(False)> _
                Public ReadOnly Property FieldErrorMessage() As String
            Get
                Return m_ErrorMessage
            End Get
        End Property

        Private Function IsValid(ByVal passtext As String) As Boolean
            ' If we don't allow null values and the current value is null
            ' return false, or if we do allow null values and the current
            ' value is null, retrurn true

            m_ErrorMessage = ""

            If AllowNullValue = False And passtext.Trim = "" Then
                m_ErrorMessage = FieldType & " can not be null."
                Return False
            Else
                If AllowNullValue And passtext.Trim = "" Then
                    Return True
                End If
            End If

            If FieldType = FieldTypeEnum.BPDate Then
                If IsDate(Text) Then
                    Return True
                Else
                    m_ErrorMessage = FieldEnumName & " is not valid."
                    Return False
                End If
            End If

            ' If the value is only a decimal or the negative sign, then it is invalid
            If Not IsNumeric(passtext) Then
                m_ErrorMessage = FieldEnumName & " must be numeric."
                Return False
            End If

            ' If we do not allow zero values and the entered value is zero
            ' then return an error
            If AllowZeroValue = False And CDec(passtext) = 0 Then
                m_ErrorMessage = FieldEnumName & " can not be 0."
                Return False
            End If

            ' If we need to check the greater than field
            m_ErrorMessage = FieldEnumName & " failed greater than test."
            If CheckGreaterThan Then
                ' If the CheckGreaterThan Object has not been set, return false
                If CheckGreaterThanObject Is Nothing Then
                    Return False
                Else
                    ' If the CheckGreaterThan object is not numeric, return false
                    If Not IsNumeric(CheckGreaterThanObject.Text) Then
                        Return False
                    Else
                        ' If the checkGreaterThan objext is less than or equal to this objects
                        ' text value, then return false
                        If CDec(CheckGreaterThanObject.Text) <= CDec(Me.Text) Then
                            Return False
                        End If
                    End If
                End If
            End If

            ' If we need to check the less than field
            m_ErrorMessage = FieldEnumName & " failed less than test."
            If CheckLessThan Then
                ' If the CheckLessThan Object has not been set, return false
                If CheckLessThanObject Is Nothing Then
                    Return False
                Else
                    ' If the CheckLessThan object is not numeric, return false
                    If Not IsNumeric(CheckLessThanObject.Text) Then
                        Return False
                    Else
                        ' If the checkGreaterThan objext is Greater than or equal to this objects
                        ' text value, then return false
                        If CDec(CheckLessThanObject.Text) >= CDec(Me.Text) Then
                            Return False
                        End If
                    End If
                End If
            End If

            ' If we made it to here we have a valid field, so return true
            m_ErrorMessage = ""
            Return True
        End Function

        Private Sub NumericTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
            Dim KeyAscii As Integer
            KeyAscii = Asc(e.KeyChar)

            Select Case KeyAscii
                Case 8, 13       ' Backspace and carriage return
                    ' it will hit this if it is a carriage return or backspace
                    ' in this case we do not want to do anything and allow it

                Case 48 To 57    ' These are the digits 0-9

                    ' If there is a decimal place already in the textbox, then check how many
                    ' decimal digits there are
                    If Not Me.Text.LastIndexOf(Chr(46)) = -1 Then
                        If Me.TextLength - Me.Text.LastIndexOf(Chr(46)) > DecimalDigits(CInt(Me.FieldType)) Then
                            'if length is over the number of decimal places for this item, don't 
                            'send message to system 
                            KeyAscii = 0
                        End If
                    End If

                Case 45                 ' Minus Sign
                    ' The number can only have one minus sign, so
                    ' if there is already one, throw this one away.
                    If InStr(Me.Text, "-") <> 0 Then
                        KeyAscii = 0
                    End If

                    ' If the insertion point is not sitting at zero
                    ' which is the beginning of the field, throw away the minus sign
                    ' because it is invalid unless it is the first character
                    If (Me.FieldType = FieldTypeEnum.BPDate Or FieldType = FieldTypeEnum.BPReelNo) Then
                        KeyAscii = 0
                    Else
                        If Me.SelectionStart <> 0 Then
                            KeyAscii = 0
                        End If
                    End If

                Case 46                 ' This is a period (decimal point)
                    If (Me.FieldType = FieldTypeEnum.BPDate Or FieldType = FieldTypeEnum.BPReelNo) Then
                        KeyAscii = 0
                    Else
                        ' If we already have a period, then throw it away
                        If InStr(Me.Text, ".") <> 0 Then
                            KeyAscii = 0
                        End If
                    End If
                Case 47                 ' This is a forward slash (/)
                    If (Me.FieldType = FieldTypeEnum.BPDate) Then
                    Else
                        KeyAscii = 0
                    End If

                Case Else
                    ' Provide no handling for other keys
                    KeyAscii = 0
            End Select

            ' If we want to throw the keystroke away, then set the event as
            ' already handled.  Otherwise let the keystroke be handled normally.
            If KeyAscii = 0 Then
                e.Handled = True
            Else
                e.Handled = False
            End If
        End Sub

        Protected Overrides Sub OnValidating(ByVal e As System.ComponentModel.CancelEventArgs)
            MyBase.OnValidating(e)
            Me.Refresh()
        End Sub

        Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
            MyBase.OnTextChanged(e)
            Me.Refresh()

            ' If we are doing the check greater than, then refresh the greater than object
            If CheckGreaterThan Then
                If Not IsNothing(CheckGreaterThanObject) Then
                    CheckGreaterThanObject.Refresh()
                End If
            End If

            ' If we are doing the check less than comparison, then refresh the less than object
            If CheckLessThan Then
                If Not IsNothing(CheckLessThanObject) Then
                    CheckLessThanObject.Refresh()
                End If
            End If
        End Sub

        Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
            MyBase.OnEnabledChanged(e)
            Me.Refresh()
        End Sub

        Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)
            Dim BackBrush As Brush
            Dim ForeBrush As Brush

            ' if the entire rectangle is zero, then call the base routine to use the
            ' system colors, and exit the routine
            If pevent.ClipRectangle.Left = 0 And pevent.ClipRectangle.Right = 0 And pevent.ClipRectangle.Height = 0 And pevent.ClipRectangle.Width = 0 Then
                MyBase.OnPaintBackground(pevent)
                Exit Sub
            End If

            If Not Enabled Then
                BackBrush = System.Drawing.SystemBrushes.Control
                ForeBrush = New SolidBrush(Color.Black)
            Else
                If Valid Then
                    If SetBackgroundColor Then
                        ' If the field is a date type, then set appropriately
                        ' otherwise check if zero value
                        If (FieldType = FieldTypeEnum.BPDate) Then
                            BackBrush = System.Drawing.SystemBrushes.Window
                        Else
                            Dim c As Decimal
                            c = CDec(Me.Text)
                            ' if this is the BPReelNo column, then make sure that the background is only gray or white
                            If (FieldType = FieldTypeEnum.BPReelNo) Then
                                If c = 0 Then
                                    BackBrush = GrayBrush(pevent.ClipRectangle)
                                Else
                                    BackBrush = New SolidBrush(Color.White)
                                End If
                            Else
                                If c = 0 Then
                                    BackBrush = GrayBrush(pevent.ClipRectangle)
                                Else
                                    BackBrush = GetBackgroundBrush(c, pevent.ClipRectangle)
                                End If
                            End If
                        End If
                        ForeBrush = New SolidBrush(Color.Black)
                    Else
                        BackBrush = System.Drawing.SystemBrushes.Window
                        ForeBrush = New SolidBrush(Color.Black)
                    End If
                Else
                    BackBrush = ErrorRedBrush(pevent.ClipRectangle)
                    ForeBrush = New SolidBrush(Color.White)
                End If
            End If

            ' Call the Fill Rectangle event
            pevent.Graphics.FillRectangle(BackBrush, pevent.ClipRectangle)
            pevent.Graphics.DrawString(Me.Text, Me.Font, ForeBrush, pevent.ClipRectangle.X, pevent.ClipRectangle.Y)
        End Sub

        Private Function GetBackgroundBrush(ByVal passValue As Decimal, ByVal bounds As Rectangle) As Brush
            Dim backBrush As Brush = Nothing

            ' This formula will come up with a factor that is either 1, .1, .01, .001, etc based on the number of decimal digits
            ' for the specific item
            Dim factor As Double = 0
            If DecimalDigits(FieldType) = 0 Then
                factor = 1
            Else
                factor = CDbl("." + "1".PadLeft(DecimalDigits(FieldType), "0"))
            End If

            ' Here we are ensuring that accurate rounding is being done for any items that were calculated by formulas
            Select Case Round(passValue, DecimalDigits(FieldType))
                Case Is < Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, DecimalDigits(FieldType))
                    backBrush = RedBrush(bounds)
                Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LRL, DecimalDigits(FieldType)) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, DecimalDigits(FieldType)) - factor
                    backBrush = YellowBrush(bounds)
                Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).LCL, DecimalDigits(FieldType)) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, DecimalDigits(FieldType)) - factor
                    backBrush = GreenBrush(bounds)
                Case Round(m_OwnerObject.QCSpecParameters(m_FieldType).UCL, DecimalDigits(FieldType)) To Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, DecimalDigits(FieldType)) - factor
                    backBrush = YellowBrush(bounds)
                Case Is >= Round(m_OwnerObject.QCSpecParameters(m_FieldType).URL, DecimalDigits(FieldType))
                    backBrush = RedBrush(bounds)
            End Select

            Return backBrush
        End Function

        Protected Overrides Sub OnGotFocus(ByVal e As System.EventArgs)
            Me.SelectionStart = 0
            Me.SelectionLength = Me.Text.Length
            m_HasFocus = True
        End Sub

        Protected Overrides Sub OnLostFocus(ByVal e As System.EventArgs)
            Me.SelectionStart = 0
            Me.SelectionLength = 0
            m_HasFocus = False
        End Sub

        Public Sub SetFocus()
            Me.Focus()
        End Sub

    End Class
End Namespace
