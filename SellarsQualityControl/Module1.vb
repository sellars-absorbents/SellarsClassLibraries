Imports System.Drawing
Imports System.Drawing.Drawing2D

Module Module1
    ' Set up the possible brushes to be used for the cell background colors
    Friend Function YellowBrush(ByVal bounds As Rectangle) As Brush
        Return New LinearGradientBrush(bounds, System.Drawing.Color.LightYellow, System.Drawing.Color.Gold, LinearGradientMode.BackwardDiagonal)
    End Function

    Friend Function ErrorRedBrush(ByVal bounds As Rectangle) As Brush
        Return New LinearGradientBrush(bounds, System.Drawing.Color.MistyRose, System.Drawing.Color.Red, LinearGradientMode.BackwardDiagonal)
    End Function

    Friend Function RedBrush(ByVal bounds As Rectangle) As Brush
        Return New LinearGradientBrush(bounds, System.Drawing.Color.MistyRose, System.Drawing.Color.DarkRed, LinearGradientMode.BackwardDiagonal)
    End Function

    Friend Function GrayBrush(ByVal bounds As Rectangle) As Brush
        Return New LinearGradientBrush(bounds, System.Drawing.Color.LightGray, System.Drawing.Color.Gray, LinearGradientMode.BackwardDiagonal)
    End Function

    Friend Function GreenBrush(ByVal bounds As Rectangle) As Brush
        Return New LinearGradientBrush(bounds, System.Drawing.Color.PaleGreen, System.Drawing.Color.DarkGreen, LinearGradientMode.BackwardDiagonal)
    End Function

    Friend Function LightRedBrush(ByVal bounds As Rectangle) As Brush
        Return New LinearGradientBrush(bounds, System.Drawing.Color.MistyRose, System.Drawing.Color.Crimson, LinearGradientMode.BackwardDiagonal)
    End Function

    Friend Function BlueBrush(ByVal bounds As Rectangle) As Brush
        Return New LinearGradientBrush(bounds, System.Drawing.Color.LightBlue, System.Drawing.Color.Blue, LinearGradientMode.BackwardDiagonal)
    End Function

End Module