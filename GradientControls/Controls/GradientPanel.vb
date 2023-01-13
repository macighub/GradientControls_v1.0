Imports System.ComponentModel
Imports System.Drawing
Imports System.Net.Mime.MediaTypeNames
Imports System.Windows.Forms.VisualStyles

Public Class GradientPanel
    Inherits Panel

    Private MouseIsOver As Boolean = False
    Private MouseIsDown As Boolean = False

    Private PanelImage As Drawing.Image

    Private SystemBackground As New Button

    Private _BackgroundStyle As BackgroundStyle
    Private _ContentAlignment As Drawing.ContentAlignment
    Private _GradientMode As Drawing2D.LinearGradientMode
    Private WithEvents _Image As ScalableImage
    Private WithEvents _InnerBorder As Border
    Private WithEvents _OuterBorder As Border
    Private _Text As String
    Private _TextHandling As TextHandling = TextHandling.None
    Private _TextImageRelation As TextImageRelation = Windows.Forms.TextImageRelation.ImageAboveText

    Public Sub New()
        MyBase.New()

        Me.DoubleBuffered = True

        MouseIsOver = False
        MouseIsDown = False

        SystemBackground.FlatStyle = FlatStyle.System
        SystemBackground.Enabled = False
        SystemBackground.Text = vbNullString
        SystemBackground.Location = New Point(0, 0)
        SystemBackground.Visible = False

        _BackgroundStyle = BackgroundStyle.Simple
        _ContentAlignment = Drawing.ContentAlignment.TopLeft
        _GradientMode = Drawing2D.LinearGradientMode.Vertical
        _Image = New ScalableImage
        _InnerBorder = New Border
        _OuterBorder = New Border
        _Text = vbNullString
        _TextHandling = TextHandling.None
        _TextImageRelation = Windows.Forms.TextImageRelation.ImageAboveText
    End Sub

    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overloads Property AutoSize() As Boolean
        Get
            Return MyBase.AutoSize
        End Get
        Set(ByVal value As Boolean)
            MyBase.AutoSize = False
        End Set
    End Property

    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overloads Property AutoSizeMode() As AutoSizeMode
        Get
            Return MyBase.AutoSize
        End Get
        Set(ByVal value As AutoSizeMode)
            MyBase.AutoSize = Nothing
        End Set
    End Property

    <DefaultValue(GetType(ImageLayout), "Stretch")>
    Public Overloads Property BackgroundImageLayout() As ImageLayout
        Get
            Return MyBase.BackgroundImageLayout
        End Get
        Set(ByVal value As ImageLayout)
            If MyBase.BackgroundImageLayout = value Then Exit Property
            MyBase.BackgroundImageLayout = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Appearance")>
    <DefaultValueAttribute(GetType(BackgroundStyle), "Simple")>
    Public Property BackgroundStyle() As BackgroundStyle
        Get
            Return _BackgroundStyle
        End Get
        Set(ByVal Value As BackgroundStyle)
            If Value = BackgroundStyle.System And Not SystemBackground.Visible Then
                SystemBackground.Size = Me.Size
                SystemBackground.Visible = True
                Me.Invalidate()
            ElseIf SystemBackground.Visible Then
                SystemBackground.Visible = False
                Me.Invalidate()
            End If
            If _BackgroundStyle = Value Then Exit Property
            _BackgroundStyle = Value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Behavior")>
    <DefaultValueAttribute(GetType(Drawing.ContentAlignment), "TopLeft")>
    Public Property ContentAlignment() As Drawing.ContentAlignment
        Get
            Return _ContentAlignment
        End Get
        Set(ByVal value As Drawing.ContentAlignment)
            If _ContentAlignment = value Then Exit Property
            _ContentAlignment = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Appearance")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property GradientMode() As Drawing2D.LinearGradientMode
        Get
            Return _GradientMode
        End Get
        Set(ByVal value As Drawing2D.LinearGradientMode)
            If _GradientMode = value Then Exit Property
            _GradientMode = value
            Me.Invalidate()
        End Set
    End Property

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("Appearance")>
    <DefaultValue(GetType(ScalableImage), "(none)")>
    Public Overloads Property Image() As ScalableImage
        Get
            Return _Image
        End Get
        Set(ByVal value As ScalableImage)
            If value Is Nothing Then
                _Image = New ScalableImage
            Else
                _Image = value
            End If
            Me.Invalidate()
        End Set
    End Property

    <Browsable(True)>
    <CategoryAttribute("Appearance")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property InnerBorder() As Border
        Get
            Return _InnerBorder
        End Get
    End Property

    <Browsable(True)>
    <CategoryAttribute("Appearance")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property OuterBorder() As Border
        Get
            Return _OuterBorder
        End Get
    End Property

    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overloads Property RightToLeft() As System.Windows.Forms.RightToLeft
        Get
            Return MyBase.RightToLeft
        End Get
        Set(ByVal value As System.Windows.Forms.RightToLeft)
            MyBase.RightToLeft = System.Windows.Forms.RightToLeft.No
        End Set
    End Property

    <Browsable(True)>
    <Bindable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    Public Overrides Property Text() As String
        Get
            Return _Text
        End Get
        Set(ByVal value As String)
            If _Text = value Then Exit Property
            _Text = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Behavior")>
    <DefaultValueAttribute(GetType(TextHandling), "AutoEllipsis")>
    Public Property TextHandling() As TextHandling
        Get
            Return _TextHandling
        End Get
        Set(ByVal value As TextHandling)
            If _TextHandling = value Then Exit Property
            _TextHandling = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Behavior")>
    <DefaultValueAttribute(GetType(TextImageRelation), "ImageAboveText")>
    Public Property TextImageRelation() As TextImageRelation
        Get
            Return _TextImageRelation
        End Get
        Set(ByVal value As TextImageRelation)
            If _TextImageRelation = value Then Exit Property
            _TextImageRelation = value
            Me.Invalidate()
        End Set
    End Property

    Private Sub GradientPanel_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        If Not MouseIsDown Then
            MouseIsDown = True
            Me.Invalidate()
        End If
    End Sub

    Private Sub GradientPanel_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
        If Not MouseIsOver Then
            MouseIsOver = True
            Me.Invalidate()
        End If
    End Sub

    Private Sub GradientPanel_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseHover
        If Not MouseIsOver Then
            MouseIsOver = True
            Me.Invalidate()
        End If
    End Sub

    Private Sub GradientPanel_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseLeave
        If MouseIsOver Then
            MouseIsOver = False
            Me.Invalidate()
        End If
    End Sub

    Private Sub GradientPanel_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        If MouseIsDown Then
            MouseIsDown = False
            Me.Invalidate()
        End If
    End Sub

    Private Sub GradientPanel_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        CreateImage()
        If PanelImage IsNot Nothing Then
            e.Graphics.DrawImageUnscaled(PanelImage, 0, 0)
        End If
    End Sub

    Private Function BackBrush(ByVal Rectangle As RectangleF) As Brush
        If _BackgroundStyle = BackgroundStyle.Gradient Then
            Dim TopColor As New Classes.ExtendedColor(BackColor)
            Dim BottomColor As New Classes.ExtendedColor(BackColor)

            If MouseIsOver AndAlso (InnerBorder.Type = Border.BorderType.Popup OrElse (InnerBorder.Type = Border.BorderType.None AndAlso OuterBorder.Type = Border.BorderType.Popup)) Then
                TopColor.Lighten(0.12)
                BottomColor.Darken(0.08)
            Else
                TopColor.Lighten(0.1)
                BottomColor.Darken(0.1)
            End If

            Return New Drawing2D.LinearGradientBrush(Rectangle, TopColor.Color, BottomColor.Color, _GradientMode)
        Else
            Return New SolidBrush(BackColor)
        End If
    End Function

    Private Sub CreateImage()
        Dim imageGraphics As Graphics
        Dim BackRect As RectangleF
        Dim BorderPadding As Padding
        Dim BorderSize As Integer = 0

        If Me.Width <= 0 Or Me.Height <= 0 Then
            Exit Sub
        End If

        PanelImage = New Bitmap(Me.Width, Me.Height)
        imageGraphics = Graphics.FromImage(PanelImage)

        imageGraphics.FillRectangle(New SolidBrush(BackColor), 0, 0, Width, Height)

        If OuterBorder.Type <> Border.BorderType.None Then
            BorderSize = OuterBorder.Size
        End If
        If InnerBorder.Type <> Border.BorderType.None Then
            BorderSize += InnerBorder.Size
        End If

        BackRect.X = BorderSize + Padding.Left
        BackRect.Y = BorderSize - 1 + Padding.Top
        BackRect.Width = MyBase.Width - 2 * BorderSize - Padding.Right - Padding.Left
        BackRect.Height = MyBase.Height - 2 * BorderSize + 1 - Padding.Bottom - Padding.Top

        If BackRect.Width > 0 AndAlso BackRect.Height > 0 Then
            imageGraphics.FillRectangle(BackBrush(BackRect), BackRect)

            If (OuterBorder.Type <> Border.BorderType.None OrElse InnerBorder.Type <> Border.BorderType.None) AndAlso MouseIsDown Then
                BackRect.X -= 1
                BackRect.Y -= 1
            End If

            If BackgroundStyle = BackgroundStyle.Image And BackgroundImage IsNot Nothing Then
                Select Case BackgroundImageLayout
                    Case ImageLayout.None
                        imageGraphics.DrawImage(BackgroundImage, BackRect.Left, BackRect.Top, BackgroundImage.Width, BackgroundImage.Height)
                    Case ImageLayout.Stretch
                        imageGraphics.DrawImage(BackgroundImage, BackRect.Left, BackRect.Top, CInt(BackRect.Width + 1), CInt(BackRect.Height + 1))
                    Case ImageLayout.Center
                        imageGraphics.DrawImage(BackgroundImage, CInt((BackRect.Width - BackgroundImage.Width) / 2) + BackRect.Left, CInt((BackRect.Height - BackgroundImage.Height) / 2) + BackRect.Top, BackgroundImage.Width, BackgroundImage.Height)
                    Case ImageLayout.Zoom
                        Dim HorizontalZoom As Double, VerticalZoom As Double

                        HorizontalZoom = BackRect.Width / BackgroundImage.Width
                        VerticalZoom = BackRect.Height / BackgroundImage.Height

                        If HorizontalZoom <= VerticalZoom Then
                            imageGraphics.DrawImage(BackgroundImage, CInt(BackRect.Left), CInt(BackRect.Top + ((BackRect.Height - (BackgroundImage.Height * HorizontalZoom)) / 2)), CInt(BackgroundImage.Width * HorizontalZoom), CInt(BackgroundImage.Height * HorizontalZoom))
                        Else
                            imageGraphics.DrawImage(BackgroundImage, CInt(BackRect.Left + ((BackRect.Width - (BackgroundImage.Width * VerticalZoom)) / 2)), BackRect.Top, CInt(BackgroundImage.Width * VerticalZoom), CInt(BackgroundImage.Height * VerticalZoom))
                        End If
                    Case ImageLayout.Tile
                        For x = BackRect.Left To BackRect.Width Step BackgroundImage.Width
                            For y = BackRect.Top To BackRect.Height Step BackgroundImage.Height
                                imageGraphics.DrawImage(BackgroundImage, x, y, BackgroundImage.Width, BackgroundImage.Height)
                            Next
                        Next
                End Select
            End If
        End If

        Dim TextArea As Rectangle = New Rectangle(0, 0, 0, 0)
        Dim TextSize As Size = New Size(0, 0)
        Dim TextRows As String()

        Select Case TextImageRelation
            Case Windows.Forms.TextImageRelation.ImageAboveText, Windows.Forms.TextImageRelation.TextAboveImage
                TextSize.Width = BackRect.Width - 10
                If _Image.BaseImage IsNot Nothing Then
                    TextSize.Height = BackRect.Height - _Image.Height - 15
                Else
                    TextSize.Height = BackRect.Height - 10
                End If
            Case Windows.Forms.TextImageRelation.ImageBeforeText, Windows.Forms.TextImageRelation.TextBeforeImage
                If _Image.BaseImage IsNot Nothing Then
                    TextSize.Width = BackRect.Width - _Image.Width - 15
                Else
                    TextSize.Width = BackRect.Width - 10
                End If
                TextSize.Height = BackRect.Height - 10
            Case Windows.Forms.TextImageRelation.Overlay
                TextSize.Width = BackRect.Width - 10
                TextSize.Height = BackRect.Height - 10
        End Select

        If TextHandling = TextHandling.None Then
            ReDim TextRows(0)
            TextRows(0) = Text
        ElseIf TextHandling = TextHandling.AutoEllipsis Then
            ReDim TextRows(0)
            TextRows(0) = TextEllipsis(TextSize.Width)
        Else
            TextRows = TextSplit(TextSize)
        End If

        TextSize = New Size(0, 0)

        For x = 0 To UBound(TextRows)
            TextSize.Height += TextRenderer.MeasureText(TextRows(x), Font).Height
            If TextRenderer.MeasureText(TextRows(x), Font).Width > TextSize.Width Then
                TextSize.Width = TextRenderer.MeasureText(TextRows(x), Font).Width
            End If
        Next

        Dim ContentSize As Size = New Size(TextSize.Width, TextSize.Height)
        Dim ImageLocation As Point, TextLocation As Point

        Select Case TextImageRelation
            Case Windows.Forms.TextImageRelation.ImageAboveText, Windows.Forms.TextImageRelation.TextAboveImage
                If _Image.BaseImage IsNot Nothing Then
                    If Image.Width > TextSize.Width Then
                        ContentSize.Width = Image.Scale.Width
                    End If
                    ContentSize.Height += Image.Scale.Height + 5
                End If
                If TextImageRelation = Windows.Forms.TextImageRelation.ImageAboveText Then
                    If _Image.BaseImage IsNot Nothing Then
                        ImageLocation = New Point(CInt((ContentSize.Width - Image.Scale.Width) / 2), 0)
                        TextLocation = New Point(CInt((ContentSize.Width - TextSize.Width) / 2), Image.Scale.Height + 5)
                    Else
                        ImageLocation = New Point(CInt(ContentSize.Width / 2), 0)
                        TextLocation = New Point(CInt((ContentSize.Width - TextSize.Width) / 2), 0)
                    End If
                Else
                    If _Image.BaseImage IsNot Nothing Then
                        ImageLocation = New Point(CInt((ContentSize.Width - Image.Scale.Width) / 2), ContentSize.Height - Image.Scale.Height)
                    Else
                        ImageLocation = New Point(CInt(ContentSize.Width / 2), ContentSize.Height)
                    End If
                    TextLocation = New Point(CInt((ContentSize.Width - TextSize.Width) / 2), 0)
                End If
            Case Windows.Forms.TextImageRelation.ImageBeforeText, Windows.Forms.TextImageRelation.TextBeforeImage
                ContentSize = New Size(TextSize.Width + 10, TextSize.Height)
                If _Image.BaseImage IsNot Nothing Then
                    If Image.Scale.Height > TextSize.Height Then
                        ContentSize.Height = Image.Scale.Height
                    End If
                    ContentSize.Width += Image.Scale.Width + 5
                End If
                If TextImageRelation = Windows.Forms.TextImageRelation.ImageBeforeText Then
                    If _Image.BaseImage IsNot Nothing Then
                        ImageLocation = New Point(0, CInt((ContentSize.Height - Image.Scale.Height) / 2))
                        TextLocation = New Point(Image.Scale.Width + 5, CInt((ContentSize.Height - TextSize.Height) / 2))
                    Else
                        ImageLocation = New Point(0, CInt(ContentSize.Height / 2))
                        TextLocation = New Point(0, CInt((ContentSize.Height - TextSize.Height) / 2))
                    End If
                Else
                    If _Image.BaseImage IsNot Nothing Then
                        ImageLocation = New Point(ContentSize.Width - Image.Scale.Width, CInt((ContentSize.Height - Image.Scale.Height) / 2))
                    Else
                        ImageLocation = New Point(ContentSize.Width, CInt(ContentSize.Width / 2))
                    End If
                    TextLocation = New Point(0, CInt((ContentSize.Height - TextSize.Height) / 2))
                End If
            Case Windows.Forms.TextImageRelation.Overlay
                If _Image.BaseImage IsNot Nothing Then
                    If Image.Scale.Width > TextSize.Width Then
                        ContentSize.Width = Image.Scale.Width
                    End If
                    If Image.Scale.Height > TextSize.Height Then
                        ContentSize.Height = Image.Scale.Height
                    End If
                    ImageLocation = New Point(CInt((ContentSize.Width - Image.Scale.Width) / 2), CInt((ContentSize.Height - Image.Scale.Height) / 2))
                Else
                    ImageLocation = New Point(CInt(ContentSize.Width / 2), CInt(ContentSize.Height / 2))
                End If
                TextLocation = New Point(CInt((ContentSize.Width - TextSize.Width) / 2), CInt((ContentSize.Height - TextSize.Height) / 2))
        End Select

        If ContentSize.Width > 0 AndAlso ContentSize.Height > 0 Then
            Dim ContentBitmap As Bitmap = New Bitmap(ContentSize.Width + 2, ContentSize.Height)
            Dim ContentGraphics As Graphics = Graphics.FromImage(ContentBitmap)
            If _Image.BaseImage IsNot Nothing Then
                ContentGraphics.DrawImage(Image, New Rectangle(ImageLocation.X, ImageLocation.Y, Image.Scale.Width, Image.Scale.Height))
            End If
            For x = 0 To UBound(TextRows)
                ContentGraphics.DrawString(TextRows(x), Font, New SolidBrush(ForeColor), TextLocation.X, TextLocation.Y)
                TextLocation.Y += TextRenderer.MeasureText(TextRows(x), Font).Height
            Next

            ContentGraphics.Save()

            Select Case ContentAlignment
                Case Drawing.ContentAlignment.TopLeft
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(BackRect.Left + 5, BackRect.Top + 5, ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.TopCenter
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(CInt(BackRect.Left + (BackRect.Width - ContentSize.Width) / 2), BackRect.Top + 5, ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.TopRight
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(BackRect.Right - 5 - ContentSize.Width, BackRect.Top + 5, ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.MiddleLeft
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(BackRect.Left + 5, CInt(BackRect.Top + (BackRect.Height - ContentSize.Height) / 2), ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.MiddleCenter
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(CInt(BackRect.Left + (BackRect.Width - ContentSize.Width) / 2), CInt(BackRect.Top + (BackRect.Height - ContentSize.Height) / 2), ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.MiddleRight
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(BackRect.Right - 5 - ContentSize.Width, CInt(BackRect.Top + (BackRect.Height - ContentSize.Height) / 2), ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.BottomLeft
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(BackRect.Left + 5, BackRect.Bottom - 5 - ContentSize.Height, ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.BottomCenter
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(CInt(BackRect.Left + (BackRect.Width - ContentSize.Width) / 2), BackRect.Bottom - 5 - ContentSize.Height, ContentBitmap.Width, ContentBitmap.Height))
                Case Drawing.ContentAlignment.BottomRight
                    imageGraphics.DrawImage(ContentBitmap, New Rectangle(BackRect.Right - 5 - ContentSize.Width, BackRect.Bottom - 5 - ContentSize.Height, ContentBitmap.Width, ContentBitmap.Height))
            End Select
        End If

        BorderPadding = New Padding(Padding.Left, Padding.Top, Padding.Right, Padding.Bottom)

        If BackgroundStyle <> BackgroundStyle.System Then
            Select Case OuterBorder.Type
                Case Border.BorderType.Flat
                    Dim FlatColor As Color

                    If MouseIsDown Then
                        FlatColor = OuterBorder.FlatStyle.MouseDownColor
                    ElseIf MouseIsOver Then
                        FlatColor = OuterBorder.FlatStyle.MouseOverColor
                    Else
                        FlatColor = OuterBorder.FlatStyle.Color
                    End If

                    For BCnt = 1 To OuterBorder.Size
                        imageGraphics.DrawRectangle(New Pen(FlatColor), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - (BorderPadding.Right + BorderPadding.Left + (BCnt) * 2) + 1, Height - (BorderPadding.Bottom + BorderPadding.Top + (BCnt) * 2) + 1)
                    Next
                Case Border.BorderType.System
                    Dim ColorTop As Classes.ExtendedColor
                    Dim ColorLeft As Classes.ExtendedColor
                    Dim ColorBottom As Classes.ExtendedColor
                    Dim ColorRight As Classes.ExtendedColor

                    If MouseIsDown Then
                        ColorTop = New Classes.ExtendedColor(SystemColors.ControlDark)
                        ColorLeft = New Classes.ExtendedColor(SystemColors.ControlDarkDark)
                        ColorBottom = New Classes.ExtendedColor(SystemColors.ControlLightLight)
                        ColorRight = New Classes.ExtendedColor(SystemColors.ControlLight)
                    Else
                        ColorTop = New Classes.ExtendedColor(SystemColors.ControlLightLight)
                        ColorLeft = New Classes.ExtendedColor(SystemColors.ControlLight)
                        ColorBottom = New Classes.ExtendedColor(SystemColors.ControlDark)
                        ColorRight = New Classes.ExtendedColor(SystemColors.ControlDarkDark)
                    End If

                    For BCnt = 1 To OuterBorder.Size
                        imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Top + BCnt - 1)
                        imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Top - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Bottom - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Top + BCnt)
                    Next
                Case Border.BorderType.Raised
                    Dim ColorTop As Classes.ExtendedColor
                    Dim ColorLeft As Classes.ExtendedColor
                    Dim ColorBottom As Classes.ExtendedColor
                    Dim ColorRight As Classes.ExtendedColor

                    If OuterBorder.Color = Color.Transparent Then
                        ColorTop = New Classes.ExtendedColor(BackColor)
                        ColorLeft = New Classes.ExtendedColor(BackColor)
                        ColorBottom = New Classes.ExtendedColor(BackColor)
                        ColorRight = New Classes.ExtendedColor(BackColor)
                    Else
                        ColorTop = New Classes.ExtendedColor(OuterBorder.Color)
                        ColorLeft = New Classes.ExtendedColor(OuterBorder.Color)
                        ColorBottom = New Classes.ExtendedColor(OuterBorder.Color)
                        ColorRight = New Classes.ExtendedColor(OuterBorder.Color)
                    End If

                    If MouseIsDown Then
                        ColorTop.Darken(0.1)
                        ColorLeft.Darken(0.12)
                        ColorBottom.Lighten(0.12)
                        ColorRight.Lighten(0.1)
                    Else
                        ColorTop.Lighten(0.12)
                        ColorLeft.Lighten(0.1)
                        ColorBottom.Darken(0.1)
                        ColorRight.Darken(0.12)
                    End If

                    For BCnt = 1 To OuterBorder.Size
                        imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt - 1)
                        imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt)
                    Next
                Case Border.BorderType.Sunken
                    Dim ColorTop As Classes.ExtendedColor
                    Dim ColorLeft As Classes.ExtendedColor
                    Dim ColorBottom As Classes.ExtendedColor
                    Dim ColorRight As Classes.ExtendedColor

                    If OuterBorder.Color = Color.Transparent Then
                        ColorTop = New Classes.ExtendedColor(BackColor)
                        ColorLeft = New Classes.ExtendedColor(BackColor)
                        ColorBottom = New Classes.ExtendedColor(BackColor)
                        ColorRight = New Classes.ExtendedColor(BackColor)
                    Else
                        ColorTop = New Classes.ExtendedColor(OuterBorder.Color)
                        ColorLeft = New Classes.ExtendedColor(OuterBorder.Color)
                        ColorBottom = New Classes.ExtendedColor(OuterBorder.Color)
                        ColorRight = New Classes.ExtendedColor(OuterBorder.Color)
                    End If

                    If MouseIsDown Then
                        ColorTop.Darken(0.14)
                        ColorLeft.Darken(0.16)
                        ColorBottom.Lighten(0.08)
                        ColorRight.Lighten(0.06)
                    Else
                        ColorTop.Darken(0.1)
                        ColorLeft.Darken(0.12)
                        ColorBottom.Lighten(0.12)
                        ColorRight.Lighten(0.1)
                    End If

                    For BCnt = 1 To OuterBorder.Size
                        imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt - 1)
                        imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt)
                    Next
                Case Border.BorderType.Popup
                    If MouseIsDown OrElse MouseIsOver Then
                        Dim ColorTop As Classes.ExtendedColor
                        Dim ColorLeft As Classes.ExtendedColor
                        Dim ColorBottom As Classes.ExtendedColor
                        Dim ColorRight As Classes.ExtendedColor

                        If OuterBorder.Color = Color.Transparent Then
                            ColorTop = New Classes.ExtendedColor(BackColor)
                            ColorLeft = New Classes.ExtendedColor(BackColor)
                            ColorBottom = New Classes.ExtendedColor(BackColor)
                            ColorRight = New Classes.ExtendedColor(BackColor)
                        Else
                            ColorTop = New Classes.ExtendedColor(OuterBorder.Color)
                            ColorLeft = New Classes.ExtendedColor(OuterBorder.Color)
                            ColorBottom = New Classes.ExtendedColor(OuterBorder.Color)
                            ColorRight = New Classes.ExtendedColor(OuterBorder.Color)
                        End If

                        If MouseIsDown Then
                            ColorTop.Darken(0.1)
                            ColorLeft.Darken(0.12)
                            ColorBottom.Lighten(0.12)
                            ColorRight.Lighten(0.1)
                        ElseIf MouseIsOver Then
                            ColorTop.Lighten(0.12)
                            ColorLeft.Lighten(0.1)
                            ColorBottom.Darken(0.1)
                            ColorRight.Darken(0.12)
                        End If

                        For BCnt = 1 To OuterBorder.Size
                            imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt - 1)
                            imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Right - BCnt)
                            imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Right - BCnt)
                            imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt)
                        Next
                    End If
            End Select

            If OuterBorder.Type <> Border.BorderType.None Then
                BorderPadding.Left += OuterBorder.Size
                BorderPadding.Top += OuterBorder.Size
                BorderPadding.Right += OuterBorder.Size
                BorderPadding.Bottom += OuterBorder.Size
            End If

            Select Case InnerBorder.Type
                Case Border.BorderType.Flat
                    Dim FlatColor As Color

                    If MouseIsDown Then
                        FlatColor = InnerBorder.FlatStyle.MouseDownColor
                    ElseIf MouseIsOver Then
                        FlatColor = InnerBorder.FlatStyle.MouseOverColor
                    Else
                        FlatColor = InnerBorder.FlatStyle.Color
                    End If

                    For BCnt = 1 To InnerBorder.Size
                        imageGraphics.DrawRectangle(New Pen(FlatColor), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - (BorderPadding.Right + BorderPadding.Left + (BCnt) * 2) + 1, Height - (BorderPadding.Bottom + BorderPadding.Top + (BCnt) * 2) + 1)
                    Next
                Case Border.BorderType.System
                    Dim ColorTop As Classes.ExtendedColor
                    Dim ColorLeft As Classes.ExtendedColor
                    Dim ColorBottom As Classes.ExtendedColor
                    Dim ColorRight As Classes.ExtendedColor

                    If MouseIsDown Then
                        ColorTop = New Classes.ExtendedColor(SystemColors.ControlDark)
                        ColorLeft = New Classes.ExtendedColor(SystemColors.ControlDarkDark)
                        ColorBottom = New Classes.ExtendedColor(SystemColors.ControlLightLight)
                        ColorRight = New Classes.ExtendedColor(SystemColors.ControlLight)
                    Else
                        ColorTop = New Classes.ExtendedColor(SystemColors.ControlLightLight)
                        ColorLeft = New Classes.ExtendedColor(SystemColors.ControlLight)
                        ColorBottom = New Classes.ExtendedColor(SystemColors.ControlDark)
                        ColorRight = New Classes.ExtendedColor(SystemColors.ControlDarkDark)
                    End If

                    For BCnt = 1 To InnerBorder.Size
                        imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt - 1)
                        imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt)
                    Next
                Case Border.BorderType.Raised
                    Dim ColorTop As Classes.ExtendedColor
                    Dim ColorLeft As Classes.ExtendedColor
                    Dim ColorBottom As Classes.ExtendedColor
                    Dim ColorRight As Classes.ExtendedColor

                    If InnerBorder.Color = Color.Transparent Then
                        ColorTop = New Classes.ExtendedColor(BackColor)
                        ColorLeft = New Classes.ExtendedColor(BackColor)
                        ColorBottom = New Classes.ExtendedColor(BackColor)
                        ColorRight = New Classes.ExtendedColor(BackColor)
                    Else
                        ColorTop = New Classes.ExtendedColor(InnerBorder.Color)
                        ColorLeft = New Classes.ExtendedColor(InnerBorder.Color)
                        ColorBottom = New Classes.ExtendedColor(InnerBorder.Color)
                        ColorRight = New Classes.ExtendedColor(InnerBorder.Color)
                    End If

                    If MouseIsDown Then
                        ColorTop.Darken(0.1)
                        ColorLeft.Darken(0.12)
                        ColorBottom.Lighten(0.12)
                        ColorRight.Lighten(0.1)
                    Else
                        ColorTop.Lighten(0.12)
                        ColorLeft.Lighten(0.1)
                        ColorBottom.Darken(0.1)
                        ColorRight.Darken(0.12)
                    End If

                    For BCnt = 1 To InnerBorder.Size
                        imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt - 1)
                        imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt)
                    Next
                Case Border.BorderType.Sunken
                    Dim ColorTop As Classes.ExtendedColor
                    Dim ColorLeft As Classes.ExtendedColor
                    Dim ColorBottom As Classes.ExtendedColor
                    Dim ColorRight As Classes.ExtendedColor

                    If InnerBorder.Color = Color.Transparent Then
                        ColorTop = New Classes.ExtendedColor(BackColor)
                        ColorLeft = New Classes.ExtendedColor(BackColor)
                        ColorBottom = New Classes.ExtendedColor(BackColor)
                        ColorRight = New Classes.ExtendedColor(BackColor)
                    Else
                        ColorTop = New Classes.ExtendedColor(InnerBorder.Color)
                        ColorLeft = New Classes.ExtendedColor(InnerBorder.Color)
                        ColorBottom = New Classes.ExtendedColor(InnerBorder.Color)
                        ColorRight = New Classes.ExtendedColor(InnerBorder.Color)
                    End If

                    If MouseIsDown Then
                        ColorTop.Darken(0.14)
                        ColorLeft.Darken(0.16)
                        ColorBottom.Lighten(0.08)
                        ColorRight.Lighten(0.06)
                    Else
                        ColorTop.Darken(0.1)
                        ColorLeft.Darken(0.12)
                        ColorBottom.Lighten(0.12)
                        ColorRight.Lighten(0.1)
                    End If

                    For BCnt = 1 To InnerBorder.Size
                        imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt - 1)
                        imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Right - BCnt)
                        imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt)
                    Next
                Case Border.BorderType.Popup
                    If MouseIsDown OrElse MouseIsOver Then
                        Dim ColorTop As Classes.ExtendedColor
                        Dim ColorLeft As Classes.ExtendedColor
                        Dim ColorBottom As Classes.ExtendedColor
                        Dim ColorRight As Classes.ExtendedColor

                        If InnerBorder.Color = Color.Transparent Then
                            ColorTop = New Classes.ExtendedColor(BackColor)
                            ColorLeft = New Classes.ExtendedColor(BackColor)
                            ColorBottom = New Classes.ExtendedColor(BackColor)
                            ColorRight = New Classes.ExtendedColor(BackColor)
                        Else
                            ColorTop = New Classes.ExtendedColor(InnerBorder.Color)
                            ColorLeft = New Classes.ExtendedColor(InnerBorder.Color)
                            ColorBottom = New Classes.ExtendedColor(InnerBorder.Color)
                            ColorRight = New Classes.ExtendedColor(InnerBorder.Color)
                        End If

                        If MouseIsDown Then
                            ColorTop.Darken(0.1)
                            ColorLeft.Darken(0.12)
                            ColorBottom.Lighten(0.12)
                            ColorRight.Lighten(0.1)
                        ElseIf MouseIsOver Then
                            ColorTop.Lighten(0.12)
                            ColorLeft.Lighten(0.1)
                            ColorBottom.Darken(0.1)
                            ColorRight.Darken(0.12)
                        End If

                        For BCnt = 1 To InnerBorder.Size
                            imageGraphics.DrawLine(New Pen(ColorTop.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt - 1)
                            imageGraphics.DrawLine(New Pen(ColorLeft.Color), BorderPadding.Left + BCnt - 1, BorderPadding.Top + BCnt - 1, BorderPadding.Left + BCnt - 1, Height - BorderPadding.Right - BCnt)
                            imageGraphics.DrawLine(New Pen(ColorBottom.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, BorderPadding.Left + BCnt, Height - BorderPadding.Right - BCnt)
                            imageGraphics.DrawLine(New Pen(ColorRight.Color), Width - BorderPadding.Right - BCnt, Height - BorderPadding.Bottom - BCnt, Width - BorderPadding.Right - BCnt, BorderPadding.Left + BCnt)
                        Next
                    End If
            End Select

            If InnerBorder.Type <> Border.BorderType.None AndAlso InnerBorder.Type <> Border.BorderType.Popup Then
                BorderPadding.Left += InnerBorder.Size
                BorderPadding.Top += InnerBorder.Size
                BorderPadding.Right += InnerBorder.Size
                BorderPadding.Bottom += InnerBorder.Size
            End If
        End If

        imageGraphics.Save()
    End Sub

    Private Function TextEllipsis(ByVal Width As Integer) As String
        Return TextEllipsis(Text, Font, Width)
    End Function

    Private Function TextEllipsis(ByVal Text As String, ByVal Font As Font, ByVal Width As Integer) As String
        Return TextFit(Text, Font, Width - TextRenderer.MeasureText("...", Font).Width) + "..."
    End Function

    Private Function TextFit(ByVal Text As String, ByVal Font As Font, ByVal Width As Integer) As String
        Dim tmpText As String

        tmpText = Text

        Do While TextRenderer.MeasureText(tmpText + " ", Font).Width > Width And Len(tmpText) > 1
            tmpText = Mid(tmpText, 1, Len(tmpText) - 1)
        Loop

        Return tmpText
    End Function

    Private Function TextSplit(ByVal DestSize As Size) As String()
        Return TextSplit(Text, Font, DestSize.Width, DestSize.Height)
    End Function

    Private Function TextSplit(ByVal Text As String, ByVal Font As Font, ByVal Width As Integer, ByVal Height As Integer) As String()
        Dim tmpResult As String()
        Dim tmpText As String
        Dim tmpHeight As Integer

        tmpText = Text

        ReDim tmpResult(0)

        Do While tmpHeight <= Height And tmpText <> vbNullString
            tmpResult(UBound(tmpResult)) = TextFit(tmpText, Font, Width)
            tmpHeight += TextRenderer.MeasureText(tmpResult(UBound(tmpResult)), Font).Height

            If tmpHeight > Height Then
                If UBound(tmpResult) > 0 Then
                    ReDim Preserve tmpResult(UBound(tmpResult) - 1)
                    tmpResult(UBound(tmpResult)) = TextEllipsis(tmpResult(UBound(tmpResult)), Font, Width)
                Else
                    tmpResult(0) = vbNullString
                End If
                tmpText = vbNullString
            Else
                If Len(tmpText) > Len(tmpResult(UBound(tmpResult))) Then
                    ReDim Preserve tmpResult(UBound(tmpResult) + 1)
                    tmpText = Mid(tmpText, Len(tmpResult(UBound(tmpResult) - 1)) + 1)

                Else
                    tmpText = vbNullString
                End If
            End If
        Loop

        Return tmpResult
    End Function

    Private Sub Border_Changed() Handles _InnerBorder.Changed, _OuterBorder.Changed
        Me.Invalidate()
    End Sub

    Private Sub _Image_Changed() Handles _Image.Changed
        Me.Invalidate()
    End Sub

End Class
