Imports System.Drawing
Imports System.Text

Namespace Classes
    Public Class ExtendedColor
        Public Enum ColorComponentsSwap
            RedGreen = 0
            RedBlue
            GreenBlue
        End Enum

        Public Sub New(ByVal color As Color)
            Me.New(color.A, color.R, color.G, color.B)
        End Sub
        Public Sub New(ByVal alpha As Byte, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte)
            _alpha = alpha
            _red = red
            _green = green
            _blue = blue

            CalculateHLS()
        End Sub
        Public Sub New(ByVal hue As Single, ByVal luminance As Single, ByVal saturation As Single)
            _alpha = 255
            _hue = hue
            _luminance = luminance
            _saturation = saturation

            CalculateColor()
        End Sub
        Public Sub New(ByVal hls As ExtendedColor)
            _alpha = hls.Alpha
            _red = hls.Red
            _blue = hls.Blue
            _green = hls.Green
            _luminance = hls.Luminance
            _hue = hls.Hue
            _saturation = hls.Saturation
        End Sub

        Private _alpha As Byte, _red As Byte, _green As Byte, _blue As Byte
        Private _hue As Single, _luminance As Single, _saturation As Single

        Public Property Alpha() As Byte
            Get
                Return _alpha
            End Get
            Set(ByVal value As Byte)
                _alpha = value
            End Set
        End Property
        Public Property Red() As Byte
            Get
                Return _red
            End Get
            Set(ByVal value As Byte)
                _red = value
                CalculateHLS()
            End Set
        End Property
        Public Property Green() As Byte
            Get
                Return _green
            End Get
            Set(ByVal value As Byte)
                _green = value
                CalculateHLS()
            End Set
        End Property
        Public Property Blue() As Byte
            Get
                Return _blue
            End Get
            Set(ByVal value As Byte)
                _blue = value
                CalculateHLS()
            End Set
        End Property

        Public Property Hue() As Single
            Get
                Return _hue
            End Get
            Set(ByVal value As Single)
                _hue = value Mod 360
                CalculateColor()
            End Set
        End Property
        Public Property Luminance() As Single
            Get
                Return _luminance
            End Get
            Set(ByVal value As Single)
                _luminance = value
                If _luminance < 0.0 Then
                    _luminance = 0.0
                End If
                If _luminance > 1.0 Then
                    _luminance = 1.0
                End If

                CalculateColor()
            End Set
        End Property
        Public Property Saturation() As Single
            Get
                Return _saturation
            End Get
            Set(ByVal value As Single)
                _saturation = value
                If _saturation < 0.0 Then
                    _saturation = 0.0
                End If
                If _saturation > 1.0 Then
                    _saturation = 1.0
                End If

                CalculateColor()
            End Set
        End Property

        Public Property Color() As Color
            Get
                Dim c As Color = Color.FromArgb(_alpha, _red, _green, _blue)
                Return c
            End Get
            Set(ByVal value As Color)
                _alpha = value.A
                _red = value.R
                _green = value.G
                _blue = value.B

                CalculateHLS()
            End Set
        End Property
        Public ReadOnly Property Complementary() As Color
            Get
                Return Color.FromArgb(_alpha, 255 - _red, 255 - _green, 255 - _blue)
            End Get
        End Property

        Public Sub Lighten(ByVal value As Single)
            Me.Luminance += value
        End Sub
        Public Sub Darken(ByVal value As Single)
            Me.Luminance -= value
        End Sub

        Public Function Clone() As ExtendedColor
            Dim c As New ExtendedColor(Me)
            Return c
        End Function
        Public Sub SwapComponents(ByVal comps As ColorComponentsSwap)
            Select Case comps
                Case ColorComponentsSwap.GreenBlue
                    SwapComps(_green, _blue)
                    Exit Select
                Case ColorComponentsSwap.RedBlue
                    SwapComps(_red, _blue)
                    Exit Select
                Case ColorComponentsSwap.RedGreen
                    SwapComps(_red, _green)
                    Exit Select
            End Select
        End Sub
        Public Sub TraslateRGB(ByVal deltaR As Byte, ByVal deltaG As Byte, ByVal deltaB As Byte)
            _red += deltaR
            _green += deltaG
            _blue += deltaB
            CalculateHLS()
        End Sub
        Public Sub SetRGB(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
            _red = r
            _green = g
            _blue = b
            CalculateHLS()
        End Sub

        Private Sub SwapComps(ByRef a As Byte, ByRef b As Byte)
            Dim x As Byte = a
            a = b
            b = x
        End Sub
        Private Sub CalculateColor()
            If _saturation = 0.0 Then
                _red = CByte(Math.Truncate(_luminance * 255.0))
                _green = _red
                _blue = _red
            Else
                Dim rm1 As Single
                Dim rm2 As Single

                If _luminance <= 0.5 Then
                    rm2 = _luminance + (_luminance * _saturation)
                Else
                    rm2 = _luminance + _saturation - (_luminance * _saturation)
                End If
                rm1 = 2.0 * _luminance - rm2
                _red = ToRGB(rm1, rm2, _hue + 120.0)
                _green = ToRGB(rm1, rm2, _hue)
                _blue = ToRGB(rm1, rm2, _hue - 120.0)
            End If
        End Sub
        Private Sub CalculateHLS()
            Dim minval As Byte = Math.Min(_red, Math.Min(_green, _blue))
            Dim maxval As Byte = Math.Max(_red, Math.Max(_green, _blue))

            Dim mdiff As Single = CSng(CSng(maxval) - CSng(minval))
            Dim msum As Single = CSng(CSng(minval) + CSng(maxval))

            _luminance = msum / 510.0

            If maxval = minval Then
                _saturation = 0.0
                _hue = 0.0
            Else
                Dim rnorm As Single = (maxval - _red) / mdiff
                Dim gnorm As Single = (maxval - _green) / mdiff
                Dim bnorm As Single = (maxval - _blue) / mdiff

                _saturation = If((_luminance <= 0.5), (mdiff / msum), (mdiff / (510.0 - msum)))

                If _red = maxval Then
                    _hue = (60.0 * (6.0 + bnorm - gnorm)) Mod 360
                End If
                If _green = maxval Then
                    _hue = 60.0 * (2.0 + rnorm - bnorm)
                End If
                If _blue = maxval Then
                    _hue = 60.0 * (4.0 + gnorm - rnorm)
                End If

                If _hue > 360.0 Then
                    _hue = _hue - 360.0
                End If
            End If
        End Sub
        Private Function ToRGB(ByVal rm1 As Single, ByVal rm2 As Single, ByVal rh As Single) As Byte
            If rh > 360.0 Then
                rh -= 360.0
            ElseIf rh < 0.0 Then
                rh += 360.0
            End If

            If rh < 60.0 Then
                rm1 = rm1 + (rm2 - rm1) * rh / 60.0
            ElseIf rh < 180.0 Then
                rm1 = rm2
            ElseIf rh < 240.0 Then
                rm1 = rm1 + (rm2 - rm1) * (240.0 - rh) / 60.0
            End If

            Return CByte(Math.Truncate(rm1 * 255))
        End Function

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder()
            sb.Append("{ ColorHLS: Alpha=" & _alpha.ToString())
            sb.Append(" , Red=" & _red.ToString())
            sb.Append(" , Green=" & _green.ToString())
            sb.Append(" , Blue=" & _blue.ToString())
            sb.Append(" , Hue=" & _hue.ToString())
            sb.Append(" , Luminance=" & _luminance.ToString())
            sb.Append(" , Saturation=" & _saturation.ToString())
            sb.Append(" }")

            Return sb.ToString()
        End Function
    End Class
End Namespace
