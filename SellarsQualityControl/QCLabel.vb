Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace Windows
    Public Class QCLabel
        Inherits System.Windows.Forms.Label
        Private m_BackColor As Integer

        ' This enumeration provides color
        Public Enum Color
            None = 0
            Gray
            Green
            Red
            LightRed
            Yellow
            Blue
        End Enum

        Public Sub New()

        End Sub

        Public Property BackgroundColor() As Color
            Get
                Return m_BackColor
            End Get
            Set(ByVal Value As Color)
                m_BackColor = Value
            End Set
        End Property

        Public Sub New(ByVal passColor As QCLabel.Color)
            m_BackColor = passColor
        End Sub

        Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)
            ' Set up the possible brushes to be used for the cell background colors

            Select Case BackgroundColor
                Case Color.None
                    MyBase.OnPaintBackground(pevent)
                Case Color.Green
                    pevent.Graphics.FillRectangle(GreenBrush(pevent.ClipRectangle), pevent.ClipRectangle)
                Case Color.Red
                    pevent.Graphics.FillRectangle(RedBrush(pevent.ClipRectangle), pevent.ClipRectangle)
                Case Color.Gray
                    pevent.Graphics.FillRectangle(GrayBrush(pevent.ClipRectangle), pevent.ClipRectangle)
                Case Color.Yellow
                    pevent.Graphics.FillRectangle(YellowBrush(pevent.ClipRectangle), pevent.ClipRectangle)
                Case Color.Blue
                    pevent.Graphics.FillRectangle(BlueBrush(pevent.ClipRectangle), pevent.ClipRectangle)
            End Select
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
            ' declare linear gradient brush for fill background of label
            OnPaintBackground(e)

            ' draw text on label
            Dim drawbrush As Brush = New SolidBrush(Me.ForeColor)
            Dim sf As New StringFormat

            ' align with center
            sf.Alignment = StringAlignment.Center
            sf.FormatFlags = StringFormatFlags.DirectionVertical

            ' set rectangle bound text
            ' Dim rectF As New RectangleF(0, Me.Height / 2 - Me.Height / 2, Me.Width, Me.Height)

            ' Rotate the text to be in the other direction
            Dim container As GraphicsContainer = e.Graphics.BeginContainer()
            e.Graphics.SetClip(e.ClipRectangle, CombineMode.Replace)
            e.Graphics.TranslateTransform(e.ClipRectangle.Width * 0.85, e.ClipRectangle.Height / 2)
            e.Graphics.RotateTransform(180)
            e.Graphics.DrawString(Me.Text, Me.Font, drawbrush, e.ClipRectangle.X, e.ClipRectangle.Y, sf)
            e.Graphics.EndContainer(container)

            ' print the output string
            ' e.Graphics.DrawString(Me.Text, Me.Font, drawbrush, rectF, sf)
        End Sub
    End Class
End Namespace