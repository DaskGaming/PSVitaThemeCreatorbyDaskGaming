Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes

Public Class ColorPickerControl
    Inherits UserControl

    ' ------------------------------------------------
    '  Internal HSV + Alpha state
    ' ------------------------------------------------
    Private _hue        As Double = 0      ' 0-360
    Private _sat        As Double = 1      ' 0-1
    Private _val        As Double = 1      ' 0-1
    Private _alpha      As Double = 1      ' 0-1
    Private _suppress   As Boolean = False
    Private _draggingWheel As Boolean = False
    Private _draggingSv    As Boolean = False
    Private _draggingAlpha As Boolean = False

    Private Const WHEEL_SIZE  As Integer = 260
    Private Const SV_SIZE     As Integer = 140
    Private Const RING_WIDTH  As Integer = 28

    ' ------------------------------------------------
    '  Dependency Property: ArgbValue
    ' ------------------------------------------------
    Public Shared ReadOnly ArgbValueProperty As DependencyProperty =
        DependencyProperty.Register(
            "ArgbValue", GetType(String), GetType(ColorPickerControl),
            New FrameworkPropertyMetadata("ffffffff",
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                AddressOf OnArgbValueChanged))

    Public Property ArgbValue As String
        Get
            Return CStr(GetValue(ArgbValueProperty))
        End Get
        Set(v As String)
            SetValue(ArgbValueProperty, v)
        End Set
    End Property

    ' ------------------------------------------------
    '  Dependency Property: SwatchColor
    ' ------------------------------------------------
    Public Shared ReadOnly SwatchColorProperty As DependencyProperty =
        DependencyProperty.Register("SwatchColor", GetType(Brush),
            GetType(ColorPickerControl), New PropertyMetadata(Brushes.White))

    Public Property SwatchColor As Brush
        Get
            Return CType(GetValue(SwatchColorProperty), Brush)
        End Get
        Set(v As Brush)
            SetValue(SwatchColorProperty, v)
        End Set
    End Property

    ' ------------------------------------------------
    '  Dependency Property: SwatchLabelColor
    ' ------------------------------------------------
    Public Shared ReadOnly SwatchLabelColorProperty As DependencyProperty =
        DependencyProperty.Register("SwatchLabelColor", GetType(Brush),
            GetType(ColorPickerControl), New PropertyMetadata(Brushes.White))

    Public Property SwatchLabelColor As Brush
        Get
            Return CType(GetValue(SwatchLabelColorProperty), Brush)
        End Get
        Set(v As Brush)
            SetValue(SwatchLabelColorProperty, v)
        End Set
    End Property

    ' ------------------------------------------------
    '  Event
    ' ------------------------------------------------
    Public Event ValueChanged As EventHandler

    ' ------------------------------------------------
    '  Loaded
    ' ------------------------------------------------
    Private Sub ColorPickerControl_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ParseArgbToHsv(ArgbValue)
        TxtArgb.Text = ArgbValue
        UpdateSwatchButton()
    End Sub

    ' ------------------------------------------------
    '  Popup opened — draw everything
    ' ------------------------------------------------
    Private Sub BtnSwatch_Click(sender As Object, e As RoutedEventArgs)
        ParseArgbToHsv(ArgbValue)
        ColorPopup.IsOpen = True
        ' Need to wait for layout before drawing
        Dispatcher.BeginInvoke(New Action(AddressOf DrawAll),
            System.Windows.Threading.DispatcherPriority.Loaded)
    End Sub

    Private Sub DrawAll()
        DrawWheel()
        DrawSvSquare()
        DrawAlphaSlider()
        UpdateMarkers()
        UpdatePreview()
    End Sub

    ' ------------------------------------------------
    '  DRAW HUE WHEEL
    ' ------------------------------------------------
    Private Sub DrawWheel()
        Dim size As Integer = WHEEL_SIZE
        Dim bmp As New WriteableBitmap(size, size, 96, 96, PixelFormats.Bgra32, Nothing)
        Dim pixels(size * size - 1) As Integer
        Dim cx As Double = size / 2.0
        Dim cy As Double = size / 2.0
        Dim outerR As Double = size / 2.0 - 2
        Dim innerR As Double = outerR - RING_WIDTH

        For y As Integer = 0 To size - 1
            For x As Integer = 0 To size - 1
                Dim dx As Double = x - cx
                Dim dy As Double = y - cy
                Dim dist As Double = Math.Sqrt(dx * dx + dy * dy)

                If dist >= innerR AndAlso dist <= outerR Then
                    Dim angle As Double = Math.Atan2(dy, dx) * 180.0 / Math.PI
                    If angle < 0 Then angle += 360
                    Dim c As Color = HsvToColor(angle, 1.0, 1.0, 255)
                    ' Soften edges
                    Dim fade As Double = 1.0
                    If dist < innerR + 2 Then fade = (dist - innerR) / 2.0
                    If dist > outerR - 2 Then fade = (outerR - dist) / 2.0
                    Dim a As Byte = CByte(Math.Max(0, Math.Min(255, CInt(255 * fade))))
                    pixels(y * size + x) = (CInt(a) << 24) Or (CInt(c.R) << 16) Or (CInt(c.G) << 8) Or CInt(c.B)
                End If
            Next
        Next

        bmp.WritePixels(New Int32Rect(0, 0, size, size), pixels, size * 4, 0)

        ' Remove only the image (index 0), keep HueMarker
        If WheelCanvas.Children.Count > 0 AndAlso TypeOf WheelCanvas.Children(0) Is Image Then
            WheelCanvas.Children.RemoveAt(0)
        End If
        Dim img As New Image With {.Width = size, .Height = size, .Source = bmp}
        WheelCanvas.Children.Insert(0, img)
    End Sub

    ' ------------------------------------------------
    '  DRAW SV SQUARE (Saturation/Value)
    ' ------------------------------------------------
    Private Sub DrawSvSquare()
        Dim size As Integer = SV_SIZE
        Dim bmp As New WriteableBitmap(size, size, 96, 96, PixelFormats.Bgra32, Nothing)
        Dim pixels(size * size - 1) As Integer

        For y As Integer = 0 To size - 1
            Dim v As Double = 1.0 - y / CDbl(size - 1)
            For x As Integer = 0 To size - 1
                Dim s As Double = x / CDbl(size - 1)
                Dim c As Color = HsvToColor(_hue, s, v, 255)
                pixels(y * size + x) = (CInt(255) << 24) Or (CInt(c.R) << 16) Or (CInt(c.G) << 8) Or CInt(c.B)
            Next
        Next

        bmp.WritePixels(New Int32Rect(0, 0, size, size), pixels, size * 4, 0)
        SvImage.Source = bmp
    End Sub

    ' ------------------------------------------------
    '  DRAW ALPHA SLIDER
    ' ------------------------------------------------
    Private Sub DrawAlphaSlider()
        Dim c As Color = HsvToColor(_hue, _sat, _val, 255)

        ' Checkerboard background
        Dim cbBrush As New DrawingBrush()
        cbBrush.TileMode = TileMode.Tile
        cbBrush.Viewport = New Rect(0, 0, 8, 8)
        cbBrush.ViewportUnits = BrushMappingMode.Absolute
        Dim dg As New DrawingGroup()
        Dim dc As DrawingContext = dg.Open()
        dc.DrawRectangle(Brushes.White, Nothing, New Rect(0, 0, 8, 8))
        dc.DrawRectangle(New SolidColorBrush(Color.FromRgb(180, 180, 180)), Nothing, New Rect(0, 0, 4, 4))
        dc.DrawRectangle(New SolidColorBrush(Color.FromRgb(180, 180, 180)), Nothing, New Rect(4, 4, 4, 4))
        dc.Close()
        cbBrush.Drawing = dg
        AlphaBg.Fill = cbBrush

        ' Color gradient overlay
        Dim grad As New LinearGradientBrush()
        grad.StartPoint = New Point(0, 0.5)
        grad.EndPoint = New Point(1, 0.5)
        grad.GradientStops.Add(New GradientStop(Color.FromArgb(0, c.R, c.G, c.B), 0))
        grad.GradientStops.Add(New GradientStop(Color.FromArgb(255, c.R, c.G, c.B), 1))
        AlphaGradient.Fill = grad
    End Sub

    ' ------------------------------------------------
    '  UPDATE MARKERS POSITION
    ' ------------------------------------------------
    Private Sub UpdateMarkers()
        ' Hue marker on the wheel ring
        Dim cx As Double = WHEEL_SIZE / 2.0
        Dim cy As Double = WHEEL_SIZE / 2.0
        Dim r As Double = WHEEL_SIZE / 2.0 - 2 - RING_WIDTH / 2.0
        Dim rad As Double = _hue * Math.PI / 180.0
        Dim mx As Double = cx + r * Math.Cos(rad)
        Dim my As Double = cy + r * Math.Sin(rad)
        HueMarker.Margin = New Thickness(mx - 7, my - 7, 0, 0)

        ' SV crosshair
        Dim sx As Double = _sat * (SV_SIZE - 1) - 6
        Dim sy As Double = (1 - _val) * (SV_SIZE - 1) - 6
        Canvas.SetLeft(SvMarker, sx)
        Canvas.SetTop(SvMarker, sy)

        ' Alpha thumb
        Dim alphaX As Double = _alpha * (AlphaCanvas.ActualWidth - 8)
        Canvas.SetLeft(AlphaThumb, alphaX)
    End Sub

    ' ------------------------------------------------
    '  UPDATE PREVIEW BOX + CURRENT ARGB TEXT
    ' ------------------------------------------------
    Private Sub UpdatePreview()
        Dim argb As String = HsvAlphaToArgb()
        Dim brush As SolidColorBrush = ArgbToBrush(argb)
        PreviewBox.Background = brush
        TxtCurrentArgb.Text = argb
    End Sub

    ' ------------------------------------------------
    '  WHEEL MOUSE EVENTS
    ' ------------------------------------------------
    Private Sub WheelCanvas_MouseDown(sender As Object, e As MouseButtonEventArgs)
        _draggingWheel = True
        WheelCanvas.CaptureMouse()
        HandleWheelMouse(e.GetPosition(WheelCanvas))
    End Sub

    Private Sub WheelCanvas_MouseMove(sender As Object, e As MouseEventArgs)
        If Not _draggingWheel Then Return
        HandleWheelMouse(e.GetPosition(WheelCanvas))
    End Sub

    Private Sub WheelCanvas_MouseUp(sender As Object, e As MouseButtonEventArgs)
        _draggingWheel = False
        WheelCanvas.ReleaseMouseCapture()
    End Sub

    Private Sub HandleWheelMouse(pos As Point)
        Dim cx As Double = WHEEL_SIZE / 2.0
        Dim cy As Double = WHEEL_SIZE / 2.0
        Dim dx As Double = pos.X - cx
        Dim dy As Double = pos.Y - cy
        Dim dist As Double = Math.Sqrt(dx * dx + dy * dy)
        Dim outerR As Double = WHEEL_SIZE / 2.0 - 2
        Dim innerR As Double = outerR - RING_WIDTH

        If dist >= innerR - 4 AndAlso dist <= outerR + 4 Then
            Dim angle As Double = Math.Atan2(dy, dx) * 180.0 / Math.PI
            If angle < 0 Then angle += 360
            _hue = angle
            DrawSvSquare()
            DrawAlphaSlider()
            UpdateMarkers()
            UpdatePreview()
            CommitColor()
        End If
    End Sub

    ' ------------------------------------------------
    '  SV SQUARE MOUSE EVENTS
    ' ------------------------------------------------
    Private Sub SvCanvas_MouseDown(sender As Object, e As MouseButtonEventArgs)
        _draggingSv = True
        SvCanvas.CaptureMouse()
        HandleSvMouse(e.GetPosition(SvCanvas))
    End Sub

    Private Sub SvCanvas_MouseMove(sender As Object, e As MouseEventArgs)
        If Not _draggingSv Then Return
        HandleSvMouse(e.GetPosition(SvCanvas))
    End Sub

    Private Sub SvCanvas_MouseUp(sender As Object, e As MouseButtonEventArgs)
        _draggingSv = False
        SvCanvas.ReleaseMouseCapture()
    End Sub

    Private Sub HandleSvMouse(pos As Point)
        _sat = Math.Max(0, Math.Min(1, pos.X / (SV_SIZE - 1)))
        _val = Math.Max(0, Math.Min(1, 1 - pos.Y / (SV_SIZE - 1)))
        DrawAlphaSlider()
        UpdateMarkers()
        UpdatePreview()
        CommitColor()
    End Sub

    ' ------------------------------------------------
    '  ALPHA SLIDER MOUSE EVENTS
    ' ------------------------------------------------
    Private Sub AlphaCanvas_MouseDown(sender As Object, e As MouseButtonEventArgs)
        _draggingAlpha = True
        AlphaCanvas.CaptureMouse()
        HandleAlphaMouse(e.GetPosition(AlphaCanvas))
    End Sub

    Private Sub AlphaCanvas_MouseMove(sender As Object, e As MouseEventArgs)
        If Not _draggingAlpha Then Return
        HandleAlphaMouse(e.GetPosition(AlphaCanvas))
    End Sub

    Private Sub AlphaCanvas_MouseUp(sender As Object, e As MouseButtonEventArgs)
        _draggingAlpha = False
        AlphaCanvas.ReleaseMouseCapture()
    End Sub

    Private Sub HandleAlphaMouse(pos As Point)
        Dim w As Double = AlphaCanvas.ActualWidth - 8
        If w <= 0 Then Return
        _alpha = Math.Max(0, Math.Min(1, pos.X / w))
        UpdateMarkers()
        UpdatePreview()
        CommitColor()
    End Sub

    ' ------------------------------------------------
    '  COMMIT COLOR TO ARGBVALUE
    ' ------------------------------------------------
    Private Sub CommitColor()
        Dim argb As String = HsvAlphaToArgb()
        ApplyArgb(argb)
    End Sub

    ' ------------------------------------------------
    '  MANUAL TEXT ENTRY
    ' ------------------------------------------------
    Private Sub TxtArgb_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _suppress Then Return
        Dim val As String = TxtArgb.Text.Trim().ToLower()
        If val.Length = 8 Then
            ApplyArgb(val)
        End If
    End Sub

    ' ------------------------------------------------
    '  DEPENDENCY PROPERTY CHANGED EXTERNALLY
    ' ------------------------------------------------
    Private Shared Sub OnArgbValueChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim ctrl As ColorPickerControl = CType(d, ColorPickerControl)
        Dim newVal As String = CStr(e.NewValue)
        ctrl.ParseArgbToHsv(newVal)
        ctrl._suppress = True
        If ctrl.TxtArgb IsNot Nothing Then ctrl.TxtArgb.Text = newVal
        ctrl._suppress = False
        ctrl.UpdateSwatchButton()
    End Sub

    ' ------------------------------------------------
    '  APPLY ARGB
    ' ------------------------------------------------
    Private Sub ApplyArgb(argb As String)
        _suppress = True
        ArgbValue    = argb
        TxtArgb.Text = argb
        TxtArgb.CaretIndex = argb.Length
        _suppress = False
        UpdateSwatchButton()
        If TxtCurrentArgb IsNot Nothing Then TxtCurrentArgb.Text = argb
        RaiseEvent ValueChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub UpdateSwatchButton()
        SwatchColor      = ArgbToBrush(ArgbValue)
        SwatchLabelColor = GetContrastBrush(ArgbValue)
    End Sub

    ' ------------------------------------------------
    '  HSV <-> ARGB CONVERSION
    ' ------------------------------------------------
    Private Sub ParseArgbToHsv(argb As String)
        Try
            If Not String.IsNullOrEmpty(argb) AndAlso argb.Length = 8 Then
                Dim a As Integer = Convert.ToInt32(argb.Substring(0, 2), 16)
                Dim r As Double  = Convert.ToInt32(argb.Substring(2, 2), 16) / 255.0
                Dim g As Double  = Convert.ToInt32(argb.Substring(4, 2), 16) / 255.0
                Dim b As Double  = Convert.ToInt32(argb.Substring(6, 2), 16) / 255.0
                _alpha = a / 255.0
                RgbToHsv(r, g, b, _hue, _sat, _val)
            End If
        Catch
        End Try
    End Sub

    Private Function HsvAlphaToArgb() As String
        Dim c As Color = HsvToColor(_hue, _sat, _val, 255)
        Dim aB As Byte = CByte(Math.Max(0, Math.Min(255, CInt(_alpha * 255))))
        Return aB.ToString("x2") & c.R.ToString("x2") & c.G.ToString("x2") & c.B.ToString("x2")
    End Function

    Private Shared Function HsvToColor(h As Double, s As Double, v As Double, a As Integer) As Color
        Dim r, g, b As Double
        If s = 0 Then
            r = v : g = v : b = v
        Else
            Dim sector As Integer = CInt(Math.Floor(h / 60)) Mod 6
            Dim f As Double = h / 60 - Math.Floor(h / 60)
            Dim p As Double = v * (1 - s)
            Dim q As Double = v * (1 - f * s)
            Dim t As Double = v * (1 - (1 - f) * s)
            Select Case sector
                Case 0 : r = v : g = t : b = p
                Case 1 : r = q : g = v : b = p
                Case 2 : r = p : g = v : b = t
                Case 3 : r = p : g = q : b = v
                Case 4 : r = t : g = p : b = v
                Case Else : r = v : g = p : b = q
            End Select
        End If
        Return Color.FromArgb(CByte(a),
                              CByte(Math.Max(0, Math.Min(255, CInt(r * 255)))),
                              CByte(Math.Max(0, Math.Min(255, CInt(g * 255)))),
                              CByte(Math.Max(0, Math.Min(255, CInt(b * 255)))))
    End Function

    Private Shared Sub RgbToHsv(r As Double, g As Double, b As Double,
                                 ByRef h As Double, ByRef s As Double, ByRef v As Double)
        Dim maxC As Double = Math.Max(r, Math.Max(g, b))
        Dim minC As Double = Math.Min(r, Math.Min(g, b))
        Dim delta As Double = maxC - minC
        v = maxC
        If maxC = 0 Then
            s = 0 : h = 0 : Return
        End If
        s = delta / maxC
        If delta = 0 Then
            h = 0 : Return
        End If
        If maxC = r Then
            h = 60 * (((g - b) / delta) Mod 6)
        ElseIf maxC = g Then
            h = 60 * ((b - r) / delta + 2)
        Else
            h = 60 * ((r - g) / delta + 4)
        End If
        If h < 0 Then h += 360
    End Sub

    ' ------------------------------------------------
    '  ARGB HEX -> BRUSH
    ' ------------------------------------------------
    Public Shared Function ArgbToBrush(hex As String) As SolidColorBrush
        Try
            If Not String.IsNullOrEmpty(hex) AndAlso hex.Length = 8 Then
                Dim a As Byte = Convert.ToByte(hex.Substring(0, 2), 16)
                Dim r As Byte = Convert.ToByte(hex.Substring(2, 2), 16)
                Dim g As Byte = Convert.ToByte(hex.Substring(4, 2), 16)
                Dim b As Byte = Convert.ToByte(hex.Substring(6, 2), 16)
                Return New SolidColorBrush(Color.FromArgb(a, r, g, b))
            End If
        Catch
        End Try
        Return New SolidColorBrush(Colors.Transparent)
    End Function

    Private Shared Function GetContrastBrush(hex As String) As SolidColorBrush
        Try
            If Not String.IsNullOrEmpty(hex) AndAlso hex.Length = 8 Then
                Dim r As Double = Convert.ToByte(hex.Substring(2, 2), 16) / 255.0
                Dim g As Double = Convert.ToByte(hex.Substring(4, 2), 16) / 255.0
                Dim b As Double = Convert.ToByte(hex.Substring(6, 2), 16) / 255.0
                Dim lum As Double = 0.2126 * r + 0.7152 * g + 0.0722 * b
                If lum > 0.45 Then
                    Return New SolidColorBrush(Color.FromArgb(180, 0, 0, 0))
                End If
            End If
        Catch
        End Try
        Return New SolidColorBrush(Color.FromArgb(180, 255, 255, 255))
    End Function

End Class
