Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.ComponentModel

Class IconCreatorWindow

    ' ------------------------------------------------
    '  Icon item model
    ' ------------------------------------------------
    Public Class IconEntry
        Implements INotifyPropertyChanged

        Private _preview As BitmapSource

        Public Property Label As String
        Public Property FileName As String
        Public Property SourcePath As String

        Public Property PreviewSource As BitmapSource
            Get
                Return _preview
            End Get
            Set(v As BitmapSource)
                _preview = v
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(PreviewSource)))
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    End Class

    ' ------------------------------------------------
    '  State
    ' ------------------------------------------------
    Private sourceFolder  As String = ""
    Private themeFolder   As String = ""
    Private currentEntry  As IconEntry = Nothing
    Private currentOrigBitmap As WriteableBitmap = Nothing

    ' Original background/foreground colors detected from the image
    Private origBgR As Byte = 246
    Private origBgG As Byte = 127
    Private origBgB As Byte = 0
    Private origFgR As Byte = 255
    Private origFgG As Byte = 255
    Private origFgB As Byte = 255

    Private iconEntries As New List(Of IconEntry)

    ' Known PS Vita icon filenames -> friendly labels
    Private ReadOnly iconNames As New Dictionary(Of String, String) From {
        {"icon_web.png",       "Browser"},
        {"icon_calendar.png",  "Calendar"},
        {"icon_photos.png",    "Camera / Photos"},
        {"icon_mail.png",      "Email"},
        {"icon_friends.png",   "Friends"},
        {"icon_cma.png",       "Content Manager"},
        {"icon_messages.png",  "Messages"},
        {"icon_music.png",     "Music"},
        {"icon_near.png",      "Near"},
        {"icon_parental.png",  "Parental Controls"},
        {"icon_party.png",     "Party / Voice Chat"},
        {"icon_ps3link.png",   "PS3 Link"},
        {"icon_ps4link.png",   "PS4 Link"},
        {"icon_power.png",     "Power"},
        {"icon_settings.png",  "Settings"},
        {"icon_trophies.png",  "Trophies"},
        {"icon_videos.png",    "Video"}
    }

    ' ------------------------------------------------
    '  Loaded
    ' ------------------------------------------------
    Private Sub IconCreatorWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        TxtBgToleranceVal.Text = CInt(SliderBgTolerance.Value).ToString()
        TxtFgToleranceVal.Text = CInt(SliderFgTolerance.Value).ToString()
        LoadBaseIconSet()
    End Sub

    ' ------------------------------------------------
    '  LOAD BASE ICON SET DROPDOWN
    ' ------------------------------------------------
    Private Sub LoadBaseIconSet()
        ' Look for baseIconSet folder relative to the exe
        Dim exeDir As String = AppDomain.CurrentDomain.BaseDirectory
        Dim baseDir As String = Path.Combine(exeDir, "baseIconSet")

        ' Also check project root (for dev/debug mode)
        If Not Directory.Exists(baseDir) Then
            Dim projectDir As String = exeDir
            For i As Integer = 0 To 4
                Dim candidate As String = Path.Combine(projectDir, "baseIconSet")
                If Directory.Exists(candidate) Then
                    baseDir = candidate
                    Exit For
                End If
                projectDir = Path.GetDirectoryName(projectDir)
                If String.IsNullOrEmpty(projectDir) Then Exit For
            Next
        End If

        If Not Directory.Exists(baseDir) Then
            TxtFooter.Text = "baseIconSet folder not found. Expected at: " & Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "baseIconSet")
            Return
        End If

        CmbBaseIcon.Items.Clear()
        Dim pngFiles As String() = Directory.GetFiles(baseDir, "*.png")
        Array.Sort(pngFiles)

        For Each filePath As String In pngFiles
            Dim fileName As String = Path.GetFileName(filePath)
            CmbBaseIcon.Items.Add(New ComboBoxItem With {
                .Content = fileName,
                .Tag     = filePath
            })
        Next

        If CmbBaseIcon.Items.Count > 0 Then
            CmbBaseIcon.SelectedIndex = 0
        End If

        TxtFooter.Text = CmbBaseIcon.Items.Count & " base icons found in baseIconSet folder."
    End Sub

    Private Sub CmbBaseIcon_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim item As ComboBoxItem = TryCast(CmbBaseIcon.SelectedItem, ComboBoxItem)
        If item Is Nothing Then Return

        Dim filePath As String = item.Tag?.ToString()
        If String.IsNullOrEmpty(filePath) OrElse Not File.Exists(filePath) Then Return

        ' Show small preview next to dropdown
        Try
            ImgBaseIconPreview.Source = LoadBitmapSafe(filePath)
        Catch
        End Try

        ' Create a temporary IconEntry and load it for editing
        Dim entry As New IconEntry With {
            .Label      = Path.GetFileNameWithoutExtension(filePath),
            .FileName   = Path.GetFileName(filePath),
            .SourcePath = filePath
        }
        entry.PreviewSource = LoadBitmapSafe(filePath)
        LoadIconForEditing(entry)
    End Sub

    ' ------------------------------------------------
    '  BACK TO MAIN
    ' ------------------------------------------------
    Private Sub BtnBackToMain_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    ' ------------------------------------------------
    '  SET SOURCE FOLDER
    ' ------------------------------------------------
    Private Sub BtnSetSourceFolder_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New System.Windows.Forms.FolderBrowserDialog With {
            .Description = "Select the folder containing the original icon PNG files"
        }
        If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return

        sourceFolder = dlg.SelectedPath
        TxtSourceFolder.Text = sourceFolder
        LoadIconList()
    End Sub

    Private _selectedBorder As Border = Nothing

    Private Sub LoadIconList()
        iconEntries.Clear()
        IconStackPanel.Children.Clear()
        AllIconsPreview.Children.Clear()

        For Each kvp In iconNames
            Dim filePath As String = Path.Combine(sourceFolder, kvp.Key)
            If File.Exists(filePath) Then
                Dim entry As New IconEntry With {
                    .Label      = kvp.Value,
                    .FileName   = kvp.Key,
                    .SourcePath = filePath
                }
                Try
                    entry.PreviewSource = LoadBitmapSafe(filePath)
                Catch
                End Try
                iconEntries.Add(entry)

                ' Build clickable row
                Dim rowGrid As New Grid()
                rowGrid.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(36)})
                rowGrid.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})

                Dim img As New System.Windows.Controls.Image With {
                    .Width = 28, .Height = 28,
                    .Source = entry.PreviewSource,
                    .VerticalAlignment = VerticalAlignment.Center
                }
                RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality)

                Dim lbl As New TextBlock With {
                    .Text = entry.Label,
                    .Foreground = New SolidColorBrush(Color.FromRgb(200, 200, 200)),
                    .FontSize = 11,
                    .VerticalAlignment = VerticalAlignment.Center,
                    .TextWrapping = TextWrapping.Wrap,
                    .Margin = New Thickness(6, 0, 0, 0)
                }
                System.Windows.Controls.Grid.SetColumn(lbl, 1)
                rowGrid.Children.Add(img)
                rowGrid.Children.Add(lbl)

                Dim row As New Border With {
                    .CornerRadius = New CornerRadius(6),
                    .Padding      = New Thickness(8, 7, 8, 7),
                    .Margin       = New Thickness(0, 2, 0, 0),
                    .Background   = Brushes.Transparent,
                    .Cursor       = System.Windows.Input.Cursors.Hand,
                    .Tag          = entry,
                    .Child        = rowGrid
                }
                AddHandler row.MouseLeftButtonUp, AddressOf IconRow_Click
                AddHandler row.MouseEnter, Sub(s, ev)
                    Dim b As Border = CType(s, Border)
                    If b IsNot _selectedBorder Then
                        b.Background = New SolidColorBrush(Color.FromRgb(30, 30, 30))
                    End If
                End Sub
                AddHandler row.MouseLeave, Sub(s, ev)
                    Dim b As Border = CType(s, Border)
                    If b IsNot _selectedBorder Then
                        b.Background = Brushes.Transparent
                    End If
                End Sub

                IconStackPanel.Children.Add(row)

                ' Right panel thumbnail
                Dim thumb As New System.Windows.Controls.Image With {
                    .Width = 40, .Height = 40,
                    .Source = entry.PreviewSource,
                    .Margin = New Thickness(3),
                    .ToolTip = entry.Label,
                    .Tag = entry,
                    .Cursor = System.Windows.Input.Cursors.Hand
                }
                RenderOptions.SetBitmapScalingMode(thumb, BitmapScalingMode.HighQuality)
                AllIconsPreview.Children.Add(thumb)
            End If
        Next

        TxtFooter.Text = iconEntries.Count & " icons loaded from source folder."
        UpdateSaveButtons()
    End Sub

    Private Sub IconRow_Click(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        ' Deselect previous
        If _selectedBorder IsNot Nothing Then
            _selectedBorder.Background = Brushes.Transparent
            _selectedBorder.BorderThickness = New Thickness(0)
        End If

        Dim row As Border = CType(sender, Border)
        row.Background = New SolidColorBrush(Color.FromRgb(42, 42, 42))
        row.BorderBrush = New SolidColorBrush(Color.FromRgb(255, 125, 0))
        row.BorderThickness = New Thickness(1)
        _selectedBorder = row

        Dim entry As IconEntry = TryCast(row.Tag, IconEntry)
        If entry IsNot Nothing Then LoadIconForEditing(entry)
    End Sub

    Private Sub IconListBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ' No longer used
    End Sub

    ' ------------------------------------------------
    '  LOAD ICON FOR EDITING
    ' ------------------------------------------------
    Private Sub LoadIconForEditing(entry As IconEntry)
        currentEntry = entry
        TxtEditorTitle.Text    = entry.Label
        TxtEditorFilename.Text = entry.FileName

        ' Show inverted hint for ps3link and ps4link
        Dim invertedIcons As String() = {"icon_ps3link.png", "icon_ps4link.png"}
        InvertedIconHint.Visibility = If(Array.IndexOf(invertedIcons, entry.FileName.ToLower()) >= 0,
                                         Visibility.Visible, Visibility.Collapsed)

        Try
            Dim bmp As BitmapImage = LoadBitmapSafe(entry.SourcePath)
            ImgOriginal.Source = bmp

            ' Force convert to Bgra32 so pixel ops are reliable
            Dim converted As New FormatConvertedBitmap(bmp, PixelFormats.Bgra32, Nothing, 0)
            currentOrigBitmap = New WriteableBitmap(converted)

            ' Auto-detect background color from top-left corner (BGRA order)
            Dim pixels(3) As Byte
            currentOrigBitmap.CopyPixels(New Int32Rect(2, 2, 1, 1), pixels, 4, 0)
            origBgB = pixels(0)
            origBgG = pixels(1)
            origBgR = pixels(2)

            ' Auto-detect foreground from center pixel
            Dim cx As Integer = currentOrigBitmap.PixelWidth \ 2
            Dim cy As Integer = currentOrigBitmap.PixelHeight \ 2
            currentOrigBitmap.CopyPixels(New Int32Rect(cx, cy, 1, 1), pixels, 4, 0)
            origFgB = pixels(0)
            origFgG = pixels(1)
            origFgR = pixels(2)

            ' Show detected color swatches
            DetectedBgSwatch.Background = New SolidColorBrush(Color.FromRgb(origBgR, origBgG, origBgB))
            DetectedFgSwatch.Background = New SolidColorBrush(Color.FromRgb(origFgR, origFgG, origFgB))

            ApplyRecolor()
        Catch ex As Exception
            TxtFooter.Text = "Error loading icon: " & ex.Message
        End Try

        UpdateSaveButtons()
    End Sub

    ' ------------------------------------------------
    '  APPLY RECOLOR — 2-layer approach
    ' ------------------------------------------------
    Private Sub ApplyRecolor()
        If currentOrigBitmap Is Nothing Then
            TxtFooter.Text = "ApplyRecolor: currentOrigBitmap is Nothing"
            Return
        End If
        Try
            Dim bgTol As Double = SliderBgTolerance.Value
            Dim fgTol As Double = SliderFgTolerance.Value

            Dim newBgR, newBgG, newBgB As Byte
            Dim newFgR, newFgG, newFgB As Byte
            ParsePickerColor(PickerBg.ArgbValue, newBgR, newBgG, newBgB)
            ParsePickerColor(PickerFg.ArgbValue, newFgR, newFgG, newFgB)

            Dim w As Integer = currentOrigBitmap.PixelWidth
            Dim h As Integer = currentOrigBitmap.PixelHeight
            Dim stride As Integer = w * 4
            Dim pixelData(stride * h - 1) As Byte
            currentOrigBitmap.CopyPixels(pixelData, stride, 0)

            TxtFooter.Text = $"Recoloring {w}x{h} | BG detected: ({origBgR},{origBgG},{origBgB}) FG: ({origFgR},{origFgG},{origFgB}) | New BG: ({newBgR},{newBgG},{newBgB})"

            For i As Integer = 0 To pixelData.Length - 1 Step 4
                Dim b As Byte = pixelData(i)
                Dim g As Byte = pixelData(i + 1)
                Dim r As Byte = pixelData(i + 2)
                Dim a As Byte = pixelData(i + 3)
                If a = 0 Then Continue For

                Dim distBg As Double = ColorDist(r, g, b, origBgR, origBgG, origBgB)
                Dim distFg As Double = ColorDist(r, g, b, origFgR, origFgG, origFgB)

                If distBg <= distFg AndAlso distBg < bgTol Then
                    Dim t As Double = 1.0 - (distBg / bgTol) ^ 0.5
                    pixelData(i)     = BlendByte(b, newBgB, t)
                    pixelData(i + 1) = BlendByte(g, newBgG, t)
                    pixelData(i + 2) = BlendByte(r, newBgR, t)
                ElseIf distFg < distBg AndAlso distFg < fgTol Then
                    Dim t As Double = 1.0 - (distFg / fgTol) ^ 0.5
                    pixelData(i)     = BlendByte(b, newFgB, t)
                    pixelData(i + 1) = BlendByte(g, newFgG, t)
                    pixelData(i + 2) = BlendByte(r, newFgR, t)
                End If
            Next

            Dim result As New WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, Nothing)
            result.WritePixels(New Int32Rect(0, 0, w, h), pixelData, stride, 0)
            result.Freeze()
            ImgRecolored.Source = result
            TxtFooter.Text &= " — Done."

        Catch ex As Exception
            TxtFooter.Text = "Recolor error: " & ex.GetType().Name & ": " & ex.Message
        End Try
    End Sub

    ' ------------------------------------------------
    '  HELPERS
    ' ------------------------------------------------
    Private Shared Function ColorDist(r1 As Byte, g1 As Byte, b1 As Byte,
                                      r2 As Byte, g2 As Byte, b2 As Byte) As Double
        Dim dr As Integer = CInt(r1) - r2
        Dim dg As Integer = CInt(g1) - g2
        Dim db As Integer = CInt(b1) - b2
        Return Math.Sqrt(dr * dr + dg * dg + db * db)
    End Function

    Private Shared Function BlendByte(orig As Byte, target As Byte, t As Double) As Byte
        Dim result As Double = orig + (CDbl(target) - CDbl(orig)) * t
        If result < 0 Then Return 0
        If result > 255 Then Return 255
        Return CByte(Math.Round(result))
    End Function

    Private Shared Sub ParsePickerColor(argbHex As String,
                                        ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
        Try
            r = Convert.ToByte(argbHex.Substring(2, 2), 16)
            g = Convert.ToByte(argbHex.Substring(4, 2), 16)
            b = Convert.ToByte(argbHex.Substring(6, 2), 16)
        Catch
            r = 255 : g = 255 : b = 255
        End Try
    End Sub

    Private Shared Function LoadBitmapSafe(fullPath As String) As BitmapImage
        Try
            Dim bmp As New BitmapImage()
            bmp.BeginInit()
            bmp.CacheOption = BitmapCacheOption.OnLoad
            bmp.UriSource   = New Uri(fullPath, UriKind.Absolute)
            bmp.EndInit()
            bmp.Freeze()
            Return bmp
        Catch
            Return Nothing
        End Try
    End Function

    ' ------------------------------------------------
    '  EVENTS
    ' ------------------------------------------------
    Private Sub BtnResetColors_Click(sender As Object, e As RoutedEventArgs)
        PickerBg.ArgbValue = "fff67f00"
        PickerFg.ArgbValue = "ffffffff"
        SliderBgTolerance.Value = 60
        SliderFgTolerance.Value = 60
        ApplyRecolor()
    End Sub

    Private Sub ColorPicker_Changed(sender As Object, e As EventArgs)
        ApplyRecolor()
    End Sub

    Private Sub Tolerance_Changed(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double))
        If TxtBgToleranceVal Is Nothing OrElse TxtFgToleranceVal Is Nothing Then Return
        TxtBgToleranceVal.Text = CInt(SliderBgTolerance.Value).ToString()
        TxtFgToleranceVal.Text = CInt(SliderFgTolerance.Value).ToString()
        ApplyRecolor()
    End Sub

    ' ------------------------------------------------
    '  SET THEME FOLDER
    ' ------------------------------------------------
    Private Sub BtnSetThemeFolder_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New System.Windows.Forms.FolderBrowserDialog With {
            .Description = "Select your theme folder (where icons will be saved)"
        }
        If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return
        themeFolder = dlg.SelectedPath
        TxtThemeFolderDisplay.Text = themeFolder
        TxtThemeFolderDisplay.Foreground = New SolidColorBrush(Color.FromRgb(204, 204, 204))
        UpdateSaveButtons()
    End Sub

    Private Sub UpdateSaveButtons()
        Dim canSave As Boolean = Not String.IsNullOrEmpty(themeFolder) AndAlso
                                 Directory.Exists(themeFolder)
        BtnSaveIcon.IsEnabled = canSave AndAlso currentEntry IsNot Nothing AndAlso
                                 ImgRecolored.Source IsNot Nothing
        BtnSaveAll.IsEnabled  = canSave AndAlso iconEntries.Count > 0
    End Sub

    ' ------------------------------------------------
    '  SAVE CURRENT ICON
    ' ------------------------------------------------
    Private Sub BtnSaveIcon_Click(sender As Object, e As RoutedEventArgs)
        If currentEntry Is Nothing OrElse ImgRecolored.Source Is Nothing Then Return
        SaveIconToFolder(currentEntry, CType(ImgRecolored.Source, BitmapSource))
        SetStatus("Saved: " & currentEntry.FileName, "#ff9933")
    End Sub

    ' ------------------------------------------------
    '  SAVE ALL ICONS (batch recolor with current colors)
    ' ------------------------------------------------
    Private Sub BtnSaveAll_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrEmpty(themeFolder) Then Return

        Dim saved As Integer = 0
        Dim original As IconEntry = currentEntry

        For Each entry As IconEntry In iconEntries
            Try
                LoadIconForEditing(entry)
                If ImgRecolored.Source IsNot Nothing Then
                    SaveIconToFolder(entry, CType(ImgRecolored.Source, BitmapSource))
                    saved += 1
                End If
            Catch
            End Try
        Next

        ' Restore previous selection
        If original IsNot Nothing Then LoadIconForEditing(original)

        SetStatus("Saved " & saved & " of " & iconEntries.Count & " icons.", "#ff9933")
    End Sub

    Private Sub SaveIconToFolder(entry As IconEntry, source As BitmapSource)
        Dim destPath As String = Path.Combine(themeFolder, entry.FileName)
        Dim encoder As New PngBitmapEncoder()
        encoder.Frames.Add(BitmapFrame.Create(source))
        Using fs As New FileStream(destPath, FileMode.Create)
            encoder.Save(fs)
        End Using

        ' Update sidebar preview to recolored version
        entry.PreviewSource = source
    End Sub

    Private Sub SetStatus(msg As String, hexColor As String)
        TxtSaveStatus.Text = msg
        Try
            Dim c As Color = CType(ColorConverter.ConvertFromString(hexColor), Color)
            TxtSaveStatus.Foreground = New SolidColorBrush(c)
        Catch
        End Try
        TxtFooter.Text = msg
    End Sub

End Class
