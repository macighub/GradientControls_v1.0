Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Design
Imports System.Drawing.Imaging
Imports System.Globalization
Imports System.Net.Mime.MediaTypeNames

<Editor(GetType(ScalableImage.ScalableImageEditor), GetType(System.Drawing.Design.UITypeEditor))>
<TypeConverter(GetType(ScalableImage.ScalableImageConverter))>
<Serializable()>
Public Class ScalableImage
    Inherits System.ComponentModel.Component

    Private _BaseImage As Drawing.Image
    Private _ScaledImage As Drawing.Image
    Private ImageAttributes As New ImageAttributes
    Private WithEvents _Scale As ImageScale = New ImageScale(0, 0)
    Private _TransparentKey As Color = Color.Empty

    Public Event Changed()

    Public Overloads Shared Widening Operator CType(ByVal Image As System.Drawing.Image) As ScalableImage
        Return New ScalableImage(Image)
    End Operator

    Public Overloads Shared Widening Operator CType(ByVal ScalableImage As ScalableImage) As System.Drawing.Image
        Return ScalableImage.ScaledImage
    End Operator

    <NotifyParentProperty(True)>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(GetType(System.Drawing.Image), "Nothing")>
    Public Overloads Property BaseImage() As Drawing.Image
        Get
            Return _BaseImage
        End Get
        Set(ByVal value As Drawing.Image)
            _BaseImage = value
            _ScaledImage = ScaleImage(value)
            RaiseEvent Changed()
        End Set
    End Property

    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Height() As Integer
        Get
            If BaseImage IsNot Nothing Then
                Return BaseImage.Height
            Else
                Return 0
            End If
        End Get
    End Property

    <NotifyParentProperty(True)>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue("0; 0")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Scale() As ImageScale
        Get
            Return _Scale
        End Get
        Set(ByVal value As ImageScale)
            If value.Width = _Scale.Width AndAlso value.Height = _Scale.Height AndAlso value.Enabled = _Scale.Enabled Then Exit Property
            _Scale = value
            _ScaledImage = ScaleImage(_BaseImage)
            RaiseEvent Changed()
        End Set
    End Property

    <Browsable(False)>
    <NotifyParentProperty(True)>
    <RefreshProperties(RefreshProperties.All)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DefaultValue(GetType(System.Drawing.Image), "Nothing")>
    Public ReadOnly Property ScaledImage() As Drawing.Image
        Get
            Return _ScaledImage
        End Get
    End Property

    <DefaultValue(GetType(System.Drawing.Color), "System.Drawing.Color.Empty")>
    Public Property TransparentKey() As Color
        Get
            Return _TransparentKey
        End Get
        Set(ByVal value As Color)
            _TransparentKey = value

            Dim colorLow As Classes.ExtendedColor
            Dim colorHigh As Classes.ExtendedColor

            If value = Color.Empty Then
                ImageAttributes.ClearColorKey()
            Else
                If value = Color.Transparent Then
                    Dim tmpBitmap As Bitmap = New Bitmap(_BaseImage.Width, _BaseImage.Height)
                    Dim tmpGraphics As Graphics = Graphics.FromImage(tmpBitmap)
                    tmpGraphics.DrawImage(_BaseImage, 0, 0)
                    tmpGraphics.Save()

                    colorLow = New Classes.ExtendedColor(tmpBitmap.GetPixel(0, 0))
                    colorHigh = New Classes.ExtendedColor(tmpBitmap.GetPixel(0, 0))
                Else
                    colorLow = New Classes.ExtendedColor(value)
                    colorHigh = New Classes.ExtendedColor(value)
                End If
                colorLow.Darken(0.1)
                colorHigh.Lighten(0.1)
                ImageAttributes.SetColorKey(colorLow.Color, colorHigh.Color)
            End If

            _ScaledImage = ScaleImage(_BaseImage)
            RaiseEvent Changed()
        End Set
    End Property

    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Width() As Integer
        Get
            If BaseImage IsNot Nothing Then
                Return BaseImage.Width
            Else
                Return 0
            End If
        End Get
    End Property

    Private Function ScaleImage(ByVal Image As Drawing.Image) As Drawing.Image
        If Image Is Nothing Then Return Nothing

        Dim tmpScaleSize As Size = New Size(_Scale.Width, _Scale.Height)

        If _Scale.Width = 0 Or Not Scale.Enabled Then
            tmpScaleSize.Width = _BaseImage.Width
        End If
        If _Scale.Height = 0 Or Not Scale.Enabled Then
            tmpScaleSize.Height = _BaseImage.Height
        End If

        Dim tmpBitmap As Bitmap = New Bitmap(tmpScaleSize.Width, tmpScaleSize.Height)
        Dim tmpGraphics As Graphics = Graphics.FromImage(tmpBitmap)

        tmpGraphics.DrawImage(Image, New Rectangle(0, 0, tmpScaleSize.Width, tmpScaleSize.Height), 0, 0, Image.Width, Image.Height, GraphicsUnit.Pixel, ImageAttributes)
        tmpGraphics.Save()

        Return tmpBitmap

    End Function

    Public Sub New()
        Me.New(Nothing)
    End Sub

    Public Sub New(ByVal Image As Drawing.Image)
        If Image IsNot Nothing Then
            _BaseImage = Image
            _ScaledImage = ScaleImage(Image)
            RaiseEvent Changed()
        End If
    End Sub

    Private Sub _ScaleSize_Changed() Handles _Scale.Changed
        _ScaledImage = ScaleImage(_BaseImage)
        RaiseEvent Changed()
    End Sub

    <TypeConverter(GetType(ImageScale.ImageScaleConverter))>
    Public Class ImageScale
        Private _Width As Integer
        Private _Height As Integer
        Private _Enabled As Boolean

        Public Event Changed()

        <NotifyParentProperty(True)>
        <RefreshProperties(RefreshProperties.All)>
        <DefaultValue(False)>
        Public Property Enabled() As Boolean
            Get
                Return _Enabled
            End Get
            Set(ByVal value As Boolean)
                If _Enabled = value Then Exit Property
                _Enabled = value
                RaiseEvent Changed()
            End Set
        End Property

        <NotifyParentProperty(True)>
        <RefreshProperties(RefreshProperties.All)>
        <DefaultValue(0)>
        Public Property Width() As Integer
            Get
                Return _Width
            End Get
            Set(ByVal value As Integer)
                If _Width = value Then Exit Property
                _Width = value
                RaiseEvent Changed()
            End Set
        End Property

        <NotifyParentProperty(True)>
        <RefreshProperties(RefreshProperties.All)>
        <DefaultValue(0)>
        Public Property Height() As Integer
            Get
                Return _Height
            End Get
            Set(ByVal value As Integer)
                If _Height = value Then Exit Property
                _Height = value
                RaiseEvent Changed()
            End Set
        End Property

        Public Sub New(ByVal Width As Integer, ByVal Height As Integer)
            If Width < 0 OrElse Height < 0 Then
                Throw New ArgumentOutOfRangeException
            Else
                _Width = Width
                _Height = Height
            End If
        End Sub

        Public Sub New()
            _Width = 0
            _Height = 0
        End Sub

        Public Class ImageScaleConverter
            Inherits ExpandableObjectConverter

            Public Overloads Overrides Function CanConvertTo(
                                  ByVal context As ITypeDescriptorContext,
                                      ByVal destinationType As Type) As Boolean

                If (destinationType Is GetType(ImageScale)) Then
                    Return True
                End If

                Return MyBase.CanConvertTo(context, destinationType)
            End Function

            Public Overloads Overrides Function ConvertTo(
                                  ByVal context As ITypeDescriptorContext,
                                  ByVal culture As CultureInfo,
                                  ByVal value As Object,
                                  ByVal destinationType As System.Type) _
                         As Object

                If (destinationType Is GetType(System.String) AndAlso (TypeOf value Is ImageScale)) Then
                    If value.Enabled Then
                        Return "Scaled: " + Trim(Str(value.Width)) + "; " + Trim(Str(value.Height))
                    Else
                        Return "Not Scaled"
                    End If
                End If

                Return MyBase.ConvertTo(context, culture, value, destinationType)
            End Function

            Public Overrides Function CanConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal sourceType As System.Type) As Boolean
                If sourceType Is GetType(System.String) Then
                    Return True
                End If

                Return MyBase.CanConvertFrom(context, sourceType)
            End Function

            Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object
                If TypeOf value Is System.String Then
                    Dim tmpWidth As Integer
                    Dim tmpHeight As Integer

                    Try
                        tmpWidth = Val(Mid(value, 1, InStr(value, ";") - 1))
                        tmpHeight = Val(Mid(value, InStr(value, ";") + 1))
                    Catch
                        Throw New ArgumentException
                    End Try
                End If

                Return MyBase.ConvertFrom(context, culture, value)
            End Function
        End Class
    End Class

    Public Class ScalableImageConverter
        Inherits ExpandableObjectConverter

        Public Overloads Overrides Function CanConvertTo(
                                  ByVal context As ITypeDescriptorContext,
                                      ByVal destinationType As Type) As Boolean

            If (destinationType Is GetType(ScalableImage)) Then
                Return True
            End If

            Return MyBase.CanConvertTo(context, destinationType)
        End Function

        Public Overloads Overrides Function ConvertTo(
                                  ByVal context As ITypeDescriptorContext,
                                  ByVal culture As CultureInfo,
                                  ByVal value As Object,
                                  ByVal destinationType As System.Type) _
                         As Object

            If (destinationType Is GetType(System.String) AndAlso (value Is Nothing OrElse value.Image Is Nothing)) Then
                Return "(none)"
            ElseIf (destinationType Is GetType(System.String) AndAlso (TypeOf value Is ScalableImage)) Then
                Return "(" + value.ScaledImage.ToString + ")"
            End If

            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End Function

        Public Overrides Function CanConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal sourceType As System.Type) As Boolean
            If sourceType Is GetType(System.String) Then
                Return True
            End If

            Return MyBase.CanConvertFrom(context, sourceType)
        End Function

        Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object
            If TypeOf value Is System.String AndAlso (value = vbNullString OrElse value = "(none)" OrElse value = "Nothing" OrElse value = "New ScalableImage") OrElse value Is Nothing Then
                Return New ScalableImage
            ElseIf TypeOf value Is Bitmap Then
                Return value
            End If

            Return (MyBase.ConvertFrom(context, culture, value))
        End Function

    End Class

    Public Class ScalableImageEditor
        Inherits ImageEditor

        Public Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            Return System.Drawing.Design.UITypeEditorEditStyle.Modal
        End Function

        Public Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As System.IServiceProvider, ByVal value As Object) As Object
            Dim _editedImage As Drawing.Image = MyBase.EditValue(context, provider, value.Image)
            Dim _newValue As New ScalableImage

            _newValue.Scale.Enabled = value.Scale.Enabled
            _newValue.Scale.Width = value.Scale.Width
            _newValue.Scale.Height = value.Scale.Height
            _newValue.TransparentKey = value.TransparentKey
            _newValue.BaseImage = _editedImage

            Return _newValue
        End Function

    End Class
End Class
