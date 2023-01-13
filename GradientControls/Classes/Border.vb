Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms.VisualStyles

<TypeConverter(GetType(Border.BorderConverter))>
Public Class Border
    Private _Color As Color
    Private WithEvents _FlatStyle As BorderFlatStyle
    Private _Size As Integer
    Private _Type As BorderType

    Public Event Changed()
    Public Event ColorChanged()
    Public Event FlatStyleChanged()
    Public Event SizeChanged()
    Public Event TypeChanged()

    Public Sub New()
        _Color = Drawing.Color.Transparent
        _Size = 1
        _Type = BorderType.Raised
        _FlatStyle = New BorderFlatStyle
    End Sub

    Public Property Color() As Color
        Get
            Return _Color
        End Get
        Set(ByVal value As Color)
            If _Color = value Then Exit Property
            _Color = value
            RaiseEvent ColorChanged()
            RaiseEvent Changed()
        End Set
    End Property

    <Browsable(True)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property FlatStyle() As BorderFlatStyle
        Get
            Return _FlatStyle
        End Get
    End Property

    Public Property Size() As Integer
        Get
            Return _Size
        End Get
        Set(ByVal value As Integer)
            If _Size = value Then Exit Property
            _Size = value
            RaiseEvent SizeChanged()
            RaiseEvent Changed()
        End Set
    End Property

    Public Property Type() As BorderType
        Get
            Return _Type
        End Get
        Set(ByVal value As BorderType)
            If _Type = value Then Exit Property
            _Type = value
            RaiseEvent TypeChanged()
            RaiseEvent Changed()
        End Set
    End Property

    Private Sub _FlatStyle_StyleChanged() Handles _FlatStyle.FlatStyleChanged
        If _Type = BorderType.Flat Then
            RaiseEvent FlatStyleChanged()
            RaiseEvent Changed()
        End If
    End Sub

    Friend Class BorderConverter
        Inherits ExpandableObjectConverter

        Public Overloads Overrides Function CanConvertTo(
                                  ByVal context As ITypeDescriptorContext,
                                      ByVal destinationType As Type) As Boolean
            If (destinationType Is GetType(Border)) Then
                Return True
            Else
                Dim tmpConverter As New ExpandableObjectConverter
                Return tmpConverter.CanConvertTo(context, destinationType)
            End If

        End Function

        Public Overloads Overrides Function ConvertTo(
                                  ByVal context As ITypeDescriptorContext,
                                  ByVal culture As CultureInfo,
                                  ByVal value As Object,
                                  ByVal destinationType As System.Type) _
                         As Object
            System.Diagnostics.Debug.Print("ConvertTo")
            If (destinationType Is GetType(System.String) AndAlso (TypeOf value Is Border)) Then
                Return vbNullString
            Else
                Dim tmpConverter As New ExpandableObjectConverter
                Return tmpConverter.ConvertTo(context, culture, value, destinationType)
            End If
        End Function

        Public Overrides Function CanConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal sourceType As System.Type) As Boolean
            If sourceType Is GetType(System.String) Then
                Return True
            Else
                Dim tmpsourcetype As System.Type
                tmpsourcetype = GetType(System.String)
                Dim tmpConverter As New ExpandableObjectConverter
                Return tmpConverter.CanConvertFrom(context, tmpsourcetype)
            End If
        End Function

        Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object
            Dim tmpConverter As New ExpandableObjectConverter
            If TypeOf value Is System.String AndAlso value = vbNullString Then
                Dim tmpValue As Object
                tmpValue = "MESComm_Controls.Controls.Border"
                Return tmpConverter.ConvertFrom(context, culture, tmpValue)
            Else
                Return tmpConverter.ConvertFrom(context, culture, value)
            End If
        End Function
    End Class

    <TypeConverter(GetType(Border.BorderFlatStyle.BorderFlatStyleConverter))>
    Public Class BorderFlatStyle

        Private _Color As Color = SystemColors.Control
        Private _MouseDownColor As Color
        Private _MouseOverColor As Color

        Public Event FlatStyleChanged()

        Public Sub New()
            _Color = SystemColors.Control
            _MouseDownColor = SystemColors.ControlDark
            _MouseOverColor = SystemColors.ControlLight
        End Sub

        Public Property Color() As Color
            Get
                Return _Color
            End Get
            Set(ByVal value As Color)
                If _Color = value Then Exit Property
                _Color = value
                RaiseEvent FlatStyleChanged()
            End Set
        End Property

        Public Property MouseDownColor() As Color
            Get
                Return _MouseDownColor
            End Get
            Set(ByVal value As Color)
                If _MouseDownColor = value Then Exit Property
                _MouseDownColor = value
                RaiseEvent FlatStyleChanged()
            End Set
        End Property

        Public Property MouseOverColor() As Color
            Get
                Return _MouseOverColor
            End Get
            Set(ByVal value As Color)
                If _MouseOverColor = value Then Exit Property
                _MouseOverColor = value
                RaiseEvent FlatStyleChanged()
            End Set
        End Property

        Friend Class BorderFlatStyleConverter
            Inherits ExpandableObjectConverter

            Public Overloads Overrides Function CanConvertTo(
                                      ByVal context As ITypeDescriptorContext,
                                          ByVal destinationType As Type) As Boolean
                If (destinationType Is GetType(BorderFlatStyle)) Then
                    Return True
                Else
                    Dim tmpConverter As New ExpandableObjectConverter
                    Return tmpConverter.CanConvertTo(context, destinationType)
                End If
            End Function

            Public Overloads Overrides Function ConvertTo(
                                      ByVal context As ITypeDescriptorContext,
                                      ByVal culture As CultureInfo,
                                      ByVal value As Object,
                                      ByVal destinationType As System.Type) _
                             As Object
                System.Diagnostics.Debug.Print("ConvertTo")
                If (destinationType Is GetType(System.String) AndAlso (TypeOf value Is BorderFlatStyle)) Then
                    Return vbNullString
                Else
                    Dim tmpConverter As New ExpandableObjectConverter
                    Return tmpConverter.ConvertTo(context, culture, value, destinationType)
                End If
            End Function
        End Class
    End Class

    Public Enum BorderType
        None
        Flat
        Popup
        Raised
        Sunken
        System
    End Enum
End Class
