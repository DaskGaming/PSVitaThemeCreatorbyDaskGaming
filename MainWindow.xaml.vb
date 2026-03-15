Imports System.IO
Imports System.Xml
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.ComponentModel
Imports Microsoft.Win32

Class MainWindow

    ' -- State --
    Private doc As New XmlDocument()
    Private xmlPath As String = ""
    Private themeFolder As String = ""
    Private currentStep As Integer = 0
    Private isEditMode As Boolean = False

    ' -- Step definitions (sidebar shows steps 1-6, welcome is step 0) --
    Private ReadOnly StepTitles As String() = {
        "Theme Setup", "Lockscreen", "Notification Bar",
        "Pages", "App Icons", "Home & Misc"
    }
    Private Const TOTAL_STEPS As Integer = 6

    ' --------------------------------
    '  STEP SIDEBAR MODEL
    ' --------------------------------
    Public Class StepItem
        Implements INotifyPropertyChanged

        Private _current As Boolean
        Private _done As Boolean

        Public Property Index As Integer
        Public Property Title As String

        Public Property IsCurrent As Boolean
            Get
                Return _current
            End Get
            Set(v As Boolean)
                _current = v
                Notify(NameOf(IsCurrent))
                Notify(NameOf(CircleBg))
                Notify(NameOf(CircleFg))
                Notify(NameOf(LabelColor))
                Notify(NameOf(RowBg))
                Notify(NameOf(NumberText))
            End Set
        End Property

        Public Property IsDone As Boolean
            Get
                Return _done
            End Get
            Set(v As Boolean)
                _done = v
                Notify(NameOf(IsDone))
                Notify(NameOf(CircleBg))
                Notify(NameOf(CircleFg))
                Notify(NameOf(LabelColor))
                Notify(NameOf(NumberText))
            End Set
        End Property

        Public ReadOnly Property NumberText As String
            Get
                Return If(IsDone, "done", (Index + 1).ToString())
            End Get
        End Property

        Public ReadOnly Property CircleBg As String
            Get
                If IsCurrent Then Return "#ff7d00"
                If IsDone Then Return "#4a2e00"
                Return "#303030"
            End Get
        End Property

        Public ReadOnly Property CircleFg As String
            Get
                If IsCurrent Then Return "#ffffff"
                If IsDone Then Return "#ff9933"
                Return "#686868"
            End Get
        End Property

        Public ReadOnly Property LabelColor As String
            Get
                If IsCurrent Then Return "#f0f0f0"
                If IsDone Then Return "#ff9933"
                Return "#787878"
            End Get
        End Property

        Public ReadOnly Property RowBg As String
            Get
                Return If(IsCurrent, "#303030", "Transparent")
            End Get
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub Notify(name As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
        End Sub
    End Class

    Private stepItems As New List(Of StepItem)

    ' --------------------------------
    '  ICON ITEM MODEL
    ' --------------------------------
    Public Class IconItem
        Implements INotifyPropertyChanged

        Private _value As String
        Private _thumbSource As System.Windows.Media.Imaging.BitmapImage
        Private _thumbVis As String = "Collapsed"

        Public Property Label As String
        Public Property Description As String
        Public Property XPath As String
        Public Property DefaultFilename As String

        Public Property Value As String
            Get
                Return _value
            End Get
            Set(v As String)
                _value = v
                Notify(NameOf(Value))
            End Set
        End Property

        Public Property ThumbnailSource As System.Windows.Media.Imaging.BitmapImage
            Get
                Return _thumbSource
            End Get
            Set(v As System.Windows.Media.Imaging.BitmapImage)
                _thumbSource = v
                Notify(NameOf(ThumbnailSource))
            End Set
        End Property

        Public Property ThumbVisibility As String
            Get
                Return _thumbVis
            End Get
            Set(v As String)
                _thumbVis = v
                Notify(NameOf(ThumbVisibility))
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub Notify(name As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
        End Sub
    End Class

    Private iconItems As New List(Of IconItem)

    ' --------------------------------
    '  WINDOW LOADED
    ' --------------------------------
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        BuildStepSidebar()
        BuildIconList()
        GoToStep(0)
    End Sub

    ' --------------------------------
    '  SIDEBAR
    ' --------------------------------
    Private Sub BuildStepSidebar()
        stepItems.Clear()
        For i As Integer = 0 To TOTAL_STEPS - 1
            stepItems.Add(New StepItem With {.Index = i, .Title = StepTitles(i)})
        Next
        StepList.ItemsSource = stepItems
    End Sub

    Private Sub UpdateSidebar()
        For Each s As StepItem In stepItems
            s.IsCurrent = (s.Index = currentStep)
            s.IsDone = (s.Index < currentStep)
        Next
    End Sub

    ' --------------------------------
    '  ICON LIST
    ' --------------------------------
    Private Sub BuildIconList()
        iconItems = New List(Of IconItem) From {
            New IconItem With {.Label = "Browser",            .Description = "Web Browser app icon.",                   .XPath = "//HomeProperty/m_browser/m_iconFilePath",     .Value = "", .DefaultFilename = "  (icon_web.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Calendar",           .Description = "Calendar app icon.",                      .XPath = "//HomeProperty/m_calendar/m_iconFilePath",    .Value = "", .DefaultFilename = "  (icon_calendar.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Camera / Photos",    .Description = "Camera and Photos app icon.",             .XPath = "//HomeProperty/m_camera/m_iconFilePath",      .Value = "", .DefaultFilename = "  (icon_photos.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Email",              .Description = "Email app icon.",                         .XPath = "//HomeProperty/m_email/m_iconFilePath",       .Value = "", .DefaultFilename = "  (icon_mail.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Friends",            .Description = "Friends List app icon.",                  .XPath = "//HomeProperty/m_friend/m_iconFilePath",      .Value = "", .DefaultFilename = "  (icon_friends.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Content Manager",    .Description = "Content Manager Assistant (CMA) icon.",  .XPath = "//HomeProperty/m_hostCollabo/m_iconFilePath", .Value = "", .DefaultFilename = "  (icon_cma.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Messages",           .Description = "Messaging app icon.",                     .XPath = "//HomeProperty/m_message/m_iconFilePath",     .Value = "", .DefaultFilename = "  (icon_messages.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Music",              .Description = "Music app icon.",                         .XPath = "//HomeProperty/m_music/m_iconFilePath",       .Value = "", .DefaultFilename = "  (icon_music.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Near",               .Description = "Near app icon.",                          .XPath = "//HomeProperty/m_near/m_iconFilePath",        .Value = "", .DefaultFilename = "  (icon_near.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Parental Controls",  .Description = "Parental Controls app icon.",             .XPath = "//HomeProperty/m_parental/m_iconFilePath",    .Value = "", .DefaultFilename = "  (icon_parental.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Party / Voice Chat", .Description = "Party Voice Chat app icon.",              .XPath = "//HomeProperty/m_party/m_iconFilePath",       .Value = "", .DefaultFilename = "  (icon_party.png  48x48  max 15KB)"},
            New IconItem With {.Label = "PS3 Link",           .Description = "PS3 Remote Play app icon.",               .XPath = "//HomeProperty/m_ps3Link/m_iconFilePath",     .Value = "", .DefaultFilename = "  (icon_ps3link.png  48x48  max 15KB)"},
            New IconItem With {.Label = "PS4 Link",           .Description = "PS4 Remote Play app icon.",               .XPath = "//HomeProperty/m_ps4Link/m_iconFilePath",     .Value = "", .DefaultFilename = "  (icon_ps4link.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Power",              .Description = "Power / On-Off app icon.",                .XPath = "//HomeProperty/m_power/m_iconFilePath",       .Value = "", .DefaultFilename = "  (icon_power.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Settings",           .Description = "System Settings app icon.",               .XPath = "//HomeProperty/m_settings/m_iconFilePath",    .Value = "", .DefaultFilename = "  (icon_settings.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Trophies",           .Description = "Trophy app icon.",                        .XPath = "//HomeProperty/m_trophy/m_iconFilePath",      .Value = "", .DefaultFilename = "  (icon_trophies.png  48x48  max 15KB)"},
            New IconItem With {.Label = "Video",              .Description = "Video app icon.",                         .XPath = "//HomeProperty/m_video/m_iconFilePath",        .Value = "", .DefaultFilename = "  (icon_videos.png  48x48  max 15KB)"}
        }
        IconList.ItemsSource = iconItems
    End Sub

    ' --------------------------------
    '  STEP NAVIGATION
    ' --------------------------------
    Private Sub GoToStep(stepNum As Integer)
        currentStep = stepNum
        UpdateSidebar()

        PanelWelcome.Visibility    = Visibility.Collapsed
        PanelSetup.Visibility      = Visibility.Collapsed
        PanelLockscreen.Visibility = Visibility.Collapsed
        PanelNotifBar.Visibility   = Visibility.Collapsed
        PanelPages.Visibility      = Visibility.Collapsed
        PanelIcons.Visibility      = Visibility.Collapsed
        PanelHomeMisc.Visibility   = Visibility.Collapsed
        PanelDone.Visibility       = Visibility.Collapsed

        Select Case stepNum
            Case 0
                PanelWelcome.Visibility = Visibility.Visible
            Case 1
                PanelSetup.Visibility = Visibility.Visible
            Case 2
                PanelLockscreen.Visibility = Visibility.Visible
            Case 3
                PanelNotifBar.Visibility = Visibility.Visible
            Case 4
                PanelPages.Visibility = Visibility.Visible
                If CmbPageSelect.SelectedIndex < 0 Then CmbPageSelect.SelectedIndex = 0
                LoadPageFields(CmbPageSelect.SelectedIndex)
            Case 5
                PanelIcons.Visibility = Visibility.Visible
            Case 6
                PanelHomeMisc.Visibility = Visibility.Visible
            Case 7
                PanelDone.Visibility = Visibility.Visible
        End Select

        ContentScroller.ScrollToTop()
        UpdatePreviewPanel()

        ' Hide nav buttons on welcome screen
        If stepNum = 0 Then
            BtnBack.Visibility = Visibility.Collapsed
            BtnNext.Visibility = Visibility.Collapsed
            BtnSaveNow.IsEnabled = False
            TxtStepIndicator.Text = ""
        ElseIf stepNum = 7 Then
            BtnNext.Visibility = Visibility.Collapsed
            BtnBack.Visibility = Visibility.Collapsed
            TxtStepIndicator.Text = ""
        Else
            BtnBack.Visibility = Visibility.Visible
            BtnNext.Visibility = Visibility.Visible
            BtnBack.IsEnabled  = (stepNum >= 1)
            BtnNext.IsEnabled  = True
            BtnNext.Content    = If(stepNum = 6, "Finish", "Save & Continue")
            TxtStepIndicator.Text = "Step " & stepNum.ToString() & " of " & TOTAL_STEPS.ToString()
        End If

        HideFooterError()
    End Sub

    Private Sub BtnNext_Click(sender As Object, e As RoutedEventArgs)
        If Not ValidateStep(currentStep) Then Return
        SaveCurrentStepToDoc()
        If currentStep = 6 Then
            SaveXmlFinal()
        Else
            GoToStep(currentStep + 1)
        End If
    End Sub

    Private Sub BtnBack_Click(sender As Object, e As RoutedEventArgs)
        ' Save current step data before going back
        If doc.DocumentElement IsNot Nothing Then
            SaveCurrentStepToDoc()
        End If

        Dim prevStep As Integer = currentStep - 1

        ' In edit mode, skip step 1 (setup) — nothing to go back to there
        If isEditMode AndAlso prevStep = 1 Then
            GoToStep(0)
        Else
            GoToStep(prevStep)
        End If
    End Sub

    Private Sub BtnSaveNow_Click(sender As Object, e As RoutedEventArgs)
        SaveCurrentStepToDoc()
        SaveXmlFile()
        SetSaveStatus("Saved at " & DateTime.Now.ToString("HH:mm:ss"))
    End Sub

    ' --------------------------------
    '  VALIDATION
    ' --------------------------------
    Private Function ValidateStep(stepNum As Integer) As Boolean
        HideFooterError()
        If stepNum = 1 AndAlso Not isEditMode Then
            If String.IsNullOrWhiteSpace(TxtTitle.Text) Then
                ShowFooterError("Please enter a theme name.")
                Return False
            End If
            If String.IsNullOrWhiteSpace(themeFolder) OrElse Not Directory.Exists(themeFolder) Then
                ShowFooterError("Please choose a save location.")
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub ShowFooterError(msg As String)
        TxtFooterStatus.Text = msg
        TxtFooterStatus.Visibility = Visibility.Visible
        TxtStepIndicator.Visibility = Visibility.Collapsed
    End Sub

    Private Sub HideFooterError()
        TxtFooterStatus.Visibility = Visibility.Collapsed
        TxtStepIndicator.Visibility = Visibility.Visible
    End Sub

    ' --------------------------------
    '  SAVE CURRENT STEP TO XML DOC
    ' --------------------------------
    Private Sub SaveCurrentStepToDoc()
        If currentStep = 1 Then
            If Not isEditMode Then InitialiseXmlDoc()
            Return
        End If
        Select Case currentStep
            Case 2
                SaveLockscreen()
            Case 3
                SaveNotifBar()
            Case 4
                SaveCurrentPage()
            Case 5
                SaveIcons()
            Case 6
                SaveHomeMisc()
        End Select
    End Sub

    Private Sub InitialiseXmlDoc()
        themeFolder = TxtThemeFolder.Text.Trim()
        xmlPath     = Path.Combine(themeFolder, "theme.xml")

        Dim themeName As String = TxtTitle.Text.Trim()
        Dim author    As String = If(String.IsNullOrWhiteSpace(TxtProvider.Text), "Unknown", TxtProvider.Text.Trim())
        Dim ver       As String = If(String.IsNullOrWhiteSpace(TxtContentVer.Text), "01.00", TxtContentVer.Text.Trim())

        doc = New XmlDocument()
        doc.LoadXml(BuildDefaultXml(themeName, author, ver))

        If ChkSyncLangs.IsChecked Then
            For Each lang In {"m_da", "m_de", "m_es", "m_fi", "m_fr", "m_it",
                              "m_nl", "m_no", "m_pl", "m_pt", "m_ru", "m_sv"}
                SetVal("//m_title/m_param/" & lang, themeName)
            Next
        End If

        TxtHeaderPath.Text = xmlPath
        BtnSaveNow.IsEnabled = True
        SaveXmlFile()
        SetSaveStatus("Created: theme.xml")
    End Sub

    Private Sub SaveLockscreen()
        SetVal("//StartScreenProperty/m_filePath",          TxtLsFilePath.Text)
        SetVal("//StartScreenProperty/m_dateColor",         TxtDateColor.ArgbValue)
        SetVal("//StartScreenProperty/m_dateLayout",        ComboTag(CmbDateLayout))
        SetVal("//StartScreenProperty/m_notifyBgColor",     TxtNotifyBgColor.ArgbValue)
        SetVal("//StartScreenProperty/m_notifyBorderColor", TxtNotifyBorderColor.ArgbValue)
        SetVal("//StartScreenProperty/m_notifyFontColor",   TxtNotifyFontColor.ArgbValue)
        SetVal("//m_startPreviewFilePath",                  TxtStartPreview.Text)
        SaveXmlFile()
    End Sub

    Private Sub SaveNotifBar()
        SetVal("//InfomationBarProperty/m_barColor",         TxtBarColor.ArgbValue)
        SetVal("//InfomationBarProperty/m_indicatorColor",   TxtIndicatorColor.ArgbValue)
        SetVal("//InfomationBarProperty/m_noticeFontColor",  TxtNoticeFontColor.ArgbValue)
        SetVal("//InfomationBarProperty/m_noticeGlowColor",  TxtNoticeGlowColor.ArgbValue)
        SetVal("//InfomationBarProperty/m_noNoticeFilePath", TxtNoNoticeFile.Text)
        SetVal("//InfomationBarProperty/m_newNoticeFilePath",TxtNewNoticeFile.Text)
        SaveXmlFile()
    End Sub

    Private Sub SaveCurrentPage()
        Dim idx As Integer = CmbPageSelect.SelectedIndex
        Dim nodes As XmlNodeList = doc.SelectNodes("//HomeProperty/m_bgParam/BackgroundParam")
        If nodes Is Nothing OrElse nodes.Count <= idx Then Return
        Dim n As XmlNode = nodes(idx)
        SetNodeVal(n, "m_imageFilePath",    TxtPageImage.Text)
        SetNodeVal(n, "m_thumbnailFilePath",TxtPageThumb.Text)
        SetNodeVal(n, "m_waveType",         TxtPageWave.Text)
        SetNodeVal(n, "m_fontColor",        TxtPageFontColor.ArgbValue)
        SetNodeVal(n, "m_fontShadow",       ComboTag(CmbPageShadow))
        SaveXmlFile()
    End Sub

    Private Sub SaveIcons()
        For Each item As IconItem In iconItems
            SetVal(item.XPath, item.Value)
        Next
        SaveXmlFile()
    End Sub

    Private Sub SaveHomeMisc()
        SetVal("//HomeProperty/m_basePageFilePath", TxtBasePage.Text)
        SetVal("//HomeProperty/m_curPageFilePath",  TxtCurPage.Text)
        SetVal("//HomeProperty/m_bgmFilePath",      TxtBgmFile.Text)
        SetVal("//m_homePreviewFilePath",            TxtHomePreview.Text)
        SetVal("//m_packageImageFilePath",           TxtPkgImage.Text)
        SaveXmlFile()
    End Sub

    ' --------------------------------
    '  FINISH
    ' --------------------------------
    Private Sub SaveXmlFinal()
        SaveHomeMisc()
        StripEmptyPageNodes()
        CopyReferencedFilesToThemeFolder()
        SaveXmlFile()
        TxtDonePath.Text = xmlPath
        GoToStep(7)
    End Sub

    ' Remove BackgroundParam nodes (pages 2-10) where both image and thumbnail are empty
    Private Sub StripEmptyPageNodes()
        Try
            Dim nodes As XmlNodeList = doc.SelectNodes("//HomeProperty/m_bgParam/BackgroundParam")
            If nodes Is Nothing Then Return
            For i As Integer = nodes.Count - 1 To 1 Step -1  ' keep page 1 (index 0) always
                Dim n As XmlNode = nodes(i)
                Dim img   As String = NodeVal(n, "m_imageFilePath").Trim()
                Dim thumb As String = NodeVal(n, "m_thumbnailFilePath").Trim()
                If String.IsNullOrEmpty(img) AndAlso String.IsNullOrEmpty(thumb) Then
                    n.ParentNode.RemoveChild(n)
                End If
            Next
        Catch
        End Try
    End Sub

    ' Copy every referenced file that exists outside themeFolder into themeFolder
    Private Sub CopyReferencedFilesToThemeFolder()
        If String.IsNullOrEmpty(themeFolder) OrElse Not Directory.Exists(themeFolder) Then Return

        Dim copied As Integer = 0

        ' Collect all text values from the XML that look like file paths
        Dim allNodes As XmlNodeList = doc.SelectNodes("//*")
        If allNodes Is Nothing Then Return

        For Each node As XmlNode In allNodes
            If node.NodeType <> XmlNodeType.Element Then Continue For
            Dim val As String = node.InnerText.Trim()
            If String.IsNullOrEmpty(val) Then Continue For

            ' Only process values that look like filenames (have an extension)
            If Path.GetExtension(val).Length < 2 Then Continue For

            ' If it's an absolute path outside the theme folder, copy it in
            Try
                Dim fullSrc As String = ""
                If File.Exists(val) Then
                    fullSrc = val
                ElseIf Not String.IsNullOrEmpty(themeFolder) Then
                    Dim candidate As String = Path.Combine(themeFolder, val)
                    If File.Exists(candidate) Then Continue For  ' already in folder
                End If

                If Not String.IsNullOrEmpty(fullSrc) Then
                    Dim destFileName As String = Path.GetFileName(fullSrc)
                    Dim destPath     As String = Path.Combine(themeFolder, destFileName)
                    If Not String.Equals(fullSrc, destPath, StringComparison.OrdinalIgnoreCase) Then
                        File.Copy(fullSrc, destPath, overwrite:=True)
                        ' Update the XML to use just the filename
                        node.InnerText = destFileName
                        copied += 1
                    End If
                End If
            Catch
            End Try
        Next

        If copied > 0 Then
            SetSaveStatus("Copied " & copied & " file(s) to theme folder.")
        End If
    End Sub

    Private Sub BtnOpenFolder_Click(sender As Object, e As RoutedEventArgs)
        If Directory.Exists(themeFolder) Then
            Process.Start("explorer.exe", themeFolder)
        End If
    End Sub

    Private Sub BtnExit_Click(sender As Object, e As RoutedEventArgs)
        Application.Current.Shutdown()
    End Sub

    Private Sub BtnStartOver_Click(sender As Object, e As RoutedEventArgs)
        doc              = New XmlDocument()
        xmlPath          = ""
        themeFolder      = ""
        isEditMode       = False
        TxtTitle.Text        = ""
        TxtProvider.Text     = ""
        TxtContentVer.Text   = "01.00"
        TxtThemeFolder.Text  = "No folder selected"
        TxtHeaderPath.Text   = ""
        BtnSaveNow.IsEnabled = False
        BuildIconList()
        GoToStep(0)
    End Sub

    ' --------------------------------
    '  WELCOME SCREEN HANDLERS
    ' --------------------------------
    Private Sub WelcomeImageRef_Click(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        PanelWelcome.Visibility  = Visibility.Collapsed
        PanelImageRef.Visibility = Visibility.Visible
    End Sub

    Private Sub BtnImageRefBack_Click(sender As Object, e As RoutedEventArgs)
        PanelImageRef.Visibility = Visibility.Collapsed
        PanelWelcome.Visibility  = Visibility.Visible
    End Sub

    Private Sub WelcomeIconCreator_Click(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        Dim creator As New IconCreatorWindow()
        creator.Owner = Me
        creator.Show()
    End Sub

    Private Sub WelcomeDownloadTheme_Click(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        Process.Start(New System.Diagnostics.ProcessStartInfo("https://psvt.ovh/") With {
            .UseShellExecute = True
        })
    End Sub

    Private Sub WelcomeNewTheme_Click(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        isEditMode = False
        GoToStep(1)
    End Sub

    Private Sub WelcomeEditTheme_Click(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        Dim dlg As New OpenFileDialog With {
            .Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*",
            .Title  = "Open Existing Theme XML"
        }
        If dlg.ShowDialog() <> True Then Return

        Try
            doc = New XmlDocument()
            doc.Load(dlg.FileName)
            xmlPath     = dlg.FileName
            themeFolder = IO.Path.GetDirectoryName(dlg.FileName)
            isEditMode  = True

            TxtHeaderPath.Text   = xmlPath
            BtnSaveNow.IsEnabled = True

            PopulateFieldsFromDoc()
            GoToStep(1)
        Catch ex As Exception
            MessageBox.Show("Could not open file:" & Environment.NewLine & ex.Message,
                            "Open Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub PopulateFieldsFromDoc()
        ' Setup fields
        TxtTitle.Text       = GetVal("//m_title/m_default")
        TxtProvider.Text    = GetVal("//m_provider/m_default")
        TxtContentVer.Text  = GetVal("//m_contentVer")
        TxtThemeFolder.Text = themeFolder

        ' Lockscreen
        TxtLsFilePath.Text        = GetVal("//StartScreenProperty/m_filePath")
        TxtDateColor.ArgbValue    = GetVal("//StartScreenProperty/m_dateColor")
        TxtNotifyBgColor.ArgbValue     = GetVal("//StartScreenProperty/m_notifyBgColor")
        TxtNotifyBorderColor.ArgbValue = GetVal("//StartScreenProperty/m_notifyBorderColor")
        TxtNotifyFontColor.ArgbValue   = GetVal("//StartScreenProperty/m_notifyFontColor")
        TxtStartPreview.Text      = GetVal("//m_startPreviewFilePath")
        SetComboByTag(CmbDateLayout, GetVal("//StartScreenProperty/m_dateLayout"))

        ' Notification bar
        TxtBarColor.ArgbValue        = GetVal("//InfomationBarProperty/m_barColor")
        TxtIndicatorColor.ArgbValue  = GetVal("//InfomationBarProperty/m_indicatorColor")
        TxtNoticeFontColor.ArgbValue = GetVal("//InfomationBarProperty/m_noticeFontColor")
        TxtNoticeGlowColor.ArgbValue = GetVal("//InfomationBarProperty/m_noticeGlowColor")
        TxtNoNoticeFile.Text  = GetVal("//InfomationBarProperty/m_noNoticeFilePath")
        TxtNewNoticeFile.Text = GetVal("//InfomationBarProperty/m_newNoticeFilePath")

        ' Pages - load page 1
        LoadPageFields(0)
        CmbPageSelect.SelectedIndex = 0

        ' App icons
        For Each item As IconItem In iconItems
            item.Value = GetVal(item.XPath)
            UpdateIconThumbnail(item)
        Next

        ' Home misc
        TxtBasePage.Text    = GetVal("//HomeProperty/m_basePageFilePath")
        TxtCurPage.Text     = GetVal("//HomeProperty/m_curPageFilePath")
        TxtBgmFile.Text     = GetVal("//HomeProperty/m_bgmFilePath")
        TxtHomePreview.Text = GetVal("//m_homePreviewFilePath")
        TxtPkgImage.Text    = GetVal("//m_packageImageFilePath")

        ' Refresh all named-field thumbnails
        Dim thumbFields As String() = {
            "TxtLsFilePath", "TxtStartPreview",
            "TxtNoNoticeFile", "TxtNewNoticeFile",
            "TxtBasePage", "TxtCurPage",
            "TxtHomePreview", "TxtPkgImage"
        }
        For Each fieldName As String In thumbFields
            UpdateThumbnail(fieldName, "Thumb" & fieldName.Substring(3))
        Next
        UpdatePreviewPanel()
    End Sub

    ' --------------------------------
    '  FOLDER HANDLING
    ' --------------------------------
    Private Sub BtnPickFolder_Click(sender As Object, e As RoutedEventArgs)
        ' Ask where to put the theme, then auto-create a subfolder named after the theme
        Dim dlg As New System.Windows.Forms.FolderBrowserDialog With {
            .Description         = "Choose where to create your theme folder",
            .ShowNewFolderButton = True
        }
        If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return

        Dim themeName  As String = SanitiseName(TxtTitle.Text)
        Dim folderName As String = If(String.IsNullOrEmpty(themeName), "NewTheme", themeName)
        Dim newPath    As String = Path.Combine(dlg.SelectedPath, folderName)

        Try
            If Not Directory.Exists(newPath) Then Directory.CreateDirectory(newPath)
            SetThemeFolder(newPath)
        Catch ex As Exception
            ShowFooterError("Could not create folder: " & ex.Message)
        End Try
    End Sub

    Private Sub SetThemeFolder(path As String)
        themeFolder         = path
        TxtThemeFolder.Text = path
    End Sub

    ' --------------------------------
    '  PAGE NAVIGATION
    ' --------------------------------
    Private Sub CmbAny_Changed(sender As Object, e As SelectionChangedEventArgs)
        UpdatePreviewPanel()
    End Sub

    Private Sub CmbPageWave_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim tag As String = ComboTag(CmbPageWave)
        TxtPageWave.Text = tag
    End Sub


    Private Sub CmbPageSelect_Changed(sender As Object, e As SelectionChangedEventArgs)
        If doc.DocumentElement Is Nothing Then Return
        LoadPageFields(CmbPageSelect.SelectedIndex)
        UpdatePreviewPanel()
    End Sub

    Private Sub BtnPrevPage_Click(sender As Object, e As RoutedEventArgs)
        If CmbPageSelect.SelectedIndex > 0 Then
            SaveCurrentPage()
            CmbPageSelect.SelectedIndex -= 1
        End If
    End Sub

    Private Sub BtnNextPage_Click(sender As Object, e As RoutedEventArgs)
        If CmbPageSelect.SelectedIndex < 9 Then
            SaveCurrentPage()
            CmbPageSelect.SelectedIndex += 1
        End If
    End Sub

    Private Sub LoadPageFields(index As Integer)
        If doc.DocumentElement Is Nothing Then Return
        Dim nodes As XmlNodeList = doc.SelectNodes("//HomeProperty/m_bgParam/BackgroundParam")
        If nodes Is Nothing OrElse nodes.Count <= index Then Return
        Dim n As XmlNode = nodes(index)
        TxtPageImage.Text     = NodeVal(n, "m_imageFilePath")
        TxtPageThumb.Text     = NodeVal(n, "m_thumbnailFilePath")
        TxtPageWave.Text      = NodeVal(n, "m_waveType")
        SetComboByTag(CmbPageWave, NodeVal(n, "m_waveType"))
        TxtPageFontColor.ArgbValue = NodeVal(n, "m_fontColor")
        SetComboByTag(CmbPageShadow, NodeVal(n, "m_fontShadow"))
        UpdateThumbnail("TxtPageImage",  "ThumbPageImage")
        UpdateThumbnail("TxtPageThumb",  "ThumbPageThumb")
        UpdatePreviewPanel()
    End Sub

    ' --------------------------------
    '  THUMBNAIL HELPERS
    ' --------------------------------
    Private Sub UpdateThumbnail(textBoxName As String, imageName As String)
        Dim tb As TextBox = TryCast(Me.FindName(textBoxName), TextBox)
        If tb Is Nothing Then Return
        Dim img As System.Windows.Controls.Image = TryCast(Me.FindName(imageName), System.Windows.Controls.Image)
        Dim brd As Border = TryCast(Me.FindName(imageName & "Border"), Border)
        If img Is Nothing Then Return

        Dim filePath As String = ResolveFilePath(tb.Text)
        If Not String.IsNullOrEmpty(filePath) AndAlso File.Exists(filePath) Then
            Try
                Dim bmp As System.Windows.Media.Imaging.BitmapImage = LoadBitmapSafe(filePath)
                img.Source = bmp
                If brd IsNot Nothing Then brd.Visibility = Visibility.Visible
            Catch
                img.Source = Nothing
                If brd IsNot Nothing Then brd.Visibility = Visibility.Collapsed
            End Try
        Else
            img.Source = Nothing
            If brd IsNot Nothing Then brd.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub UpdateIconThumbnail(item As IconItem)
        If item Is Nothing Then Return
        Dim filePath As String = ResolveFilePath(item.Value)
        If Not String.IsNullOrEmpty(filePath) AndAlso File.Exists(filePath) Then
            Try
                Dim bmp As System.Windows.Media.Imaging.BitmapImage = LoadBitmapSafe(filePath)
                item.ThumbnailSource = bmp
                item.ThumbVisibility = "Visible"
            Catch
                item.ThumbnailSource = Nothing
                item.ThumbVisibility = "Collapsed"
            End Try
        Else
            item.ThumbnailSource = Nothing
            item.ThumbVisibility = "Collapsed"
        End If
    End Sub

    ' Resolves a filename-only value against the theme folder
    Private Function ResolveFilePath(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        If File.Exists(value) Then Return Path.GetFullPath(value)
        If Not String.IsNullOrEmpty(themeFolder) Then
            Dim fullPath As String = Path.Combine(themeFolder, value)
            If File.Exists(fullPath) Then Return Path.GetFullPath(fullPath)
        End If
        Return ""
    End Function

    Private Shared Function LoadBitmapSafe(fullPath As String) As System.Windows.Media.Imaging.BitmapImage
        Try
            Dim bmp As New System.Windows.Media.Imaging.BitmapImage()
            bmp.BeginInit()
            bmp.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
            bmp.UriSource   = New Uri(fullPath, UriKind.Absolute)
            bmp.EndInit()
            bmp.Freeze()
            Return bmp
        Catch
            Return Nothing
        End Try
    End Function

    ' --------------------------------
    '  FILE BROWSE / IMPORT HANDLERS
    ' --------------------------------
    Private Sub BrowseFile_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim tbName As String = btn.Tag?.ToString()
        Dim tb As TextBox = TryCast(Me.FindName(tbName), TextBox)
        If tb Is Nothing Then Return
        Dim result As String = PickFile(tb.Text)
        If result IsNot Nothing Then
            tb.Text = result
            UpdateThumbnail(tbName, "Thumb" & tbName.Substring(3))
            UpdatePreviewPanel()
        End If
    End Sub

    Private Sub BrowseIconField_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim item As IconItem = TryCast(btn.Tag, IconItem)
        If item Is Nothing Then Return
        Dim result As String = PickFile(item.Value)
        If result IsNot Nothing Then
            item.Value = result
            UpdateIconThumbnail(item)
        End If
    End Sub

    Private Sub ImportFile_Click(sender As Object, e As RoutedEventArgs)
        If Not EnsureThemeFolder() Then Return
        Dim btn As Button = CType(sender, Button)
        Dim tbName As String = btn.Tag?.ToString()
        Dim tb As TextBox = TryCast(Me.FindName(tbName), TextBox)
        If tb Is Nothing Then Return
        Dim result As String = ImportFileToFolder(tb.Text)
        If result IsNot Nothing Then
            tb.Text = result
            UpdateThumbnail(tbName, "Thumb" & tbName.Substring(3))
            UpdatePreviewPanel()
        End If
    End Sub

    Private Sub ImportIconField_Click(sender As Object, e As RoutedEventArgs)
        If Not EnsureThemeFolder() Then Return
        Dim btn As Button = CType(sender, Button)
        Dim item As IconItem = TryCast(btn.Tag, IconItem)
        If item Is Nothing Then Return
        Dim result As String = ImportFileToFolder(item.Value)
        If result IsNot Nothing Then
            item.Value = result
            UpdateIconThumbnail(item)
        End If
    End Sub

    Private Sub BtnImportAllIcons_Click(sender As Object, e As RoutedEventArgs)
        If Not EnsureThemeFolder() Then Return

        Dim dlg As New System.Windows.Forms.FolderBrowserDialog With {
            .Description = "Choose the folder containing your icon PNG files"
        }
        If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return

        Dim srcFolder As String = dlg.SelectedPath
        Dim imported As Integer = 0
        Dim notFound As Integer = 0

        For Each item As IconItem In iconItems
            ' Extract just the filename from the DefaultFilename hint e.g. "(icon_web.png  48x48..."
            Dim hint As String = item.DefaultFilename
            Dim fnMatch As System.Text.RegularExpressions.Match =
                System.Text.RegularExpressions.Regex.Match(hint, "\(([^\s]+\.png)")
            If Not fnMatch.Success Then Continue For

            Dim fileName As String = fnMatch.Groups(1).Value
            Dim srcPath  As String = Path.Combine(srcFolder, fileName)

            If File.Exists(srcPath) Then
                Dim destPath As String = Path.Combine(themeFolder, fileName)
                Try
                    If Not String.Equals(srcPath, destPath, StringComparison.OrdinalIgnoreCase) Then
                        File.Copy(srcPath, destPath, overwrite:=True)
                    End If
                    item.Value = fileName
                    UpdateIconThumbnail(item)
                    imported += 1
                Catch
                    notFound += 1
                End Try
            Else
                notFound += 1
            End If
        Next

        SetSaveStatus("Imported " & imported & " icon" & If(imported = 1, "", "s") &
                      If(notFound > 0, " (" & notFound & " not found)", "") & ".")
    End Sub

    Private Function EnsureThemeFolder() As Boolean
        If String.IsNullOrEmpty(themeFolder) OrElse Not Directory.Exists(themeFolder) Then
            ShowFooterError("Set a theme folder in Step 1 before importing files.")
            Return False
        End If
        Return True
    End Function

    Private Function ImportFileToFolder(currentValue As String) As String
        Dim dlg As New OpenFileDialog With {
            .Filter = "Image & Audio Files (*.png;*.jpg;*.at9)|*.png;*.jpg;*.at9|All Files (*.*)|*.*",
            .Title  = "Import File into Theme Folder"
        }
        Try
            If Not String.IsNullOrEmpty(currentValue) AndAlso File.Exists(currentValue) Then
                dlg.InitialDirectory = Path.GetDirectoryName(currentValue)
            End If
        Catch
        End Try

        If dlg.ShowDialog() <> True Then Return Nothing

        Dim srcPath  As String = dlg.FileName
        Dim fileName As String = Path.GetFileName(srcPath)
        Dim destPath As String = Path.Combine(themeFolder, fileName)

        Try
            If Not String.Equals(srcPath, destPath, StringComparison.OrdinalIgnoreCase) Then
                File.Copy(srcPath, destPath, overwrite:=True)
            End If
            SetSaveStatus("Imported: " & fileName)
            Return fileName
        Catch ex As Exception
            ShowFooterError("Import failed: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function PickFile(currentValue As String) As String
        Dim dlg As New OpenFileDialog With {
            .Filter = "Image & Audio Files (*.png;*.jpg;*.at9)|*.png;*.jpg;*.at9|All Files (*.*)|*.*",
            .Title  = "Select File"
        }
        Try
            If Not String.IsNullOrEmpty(currentValue) AndAlso File.Exists(currentValue) Then
                dlg.InitialDirectory = Path.GetDirectoryName(currentValue)
            ElseIf Not String.IsNullOrEmpty(themeFolder) Then
                dlg.InitialDirectory = themeFolder
            End If
        Catch
        End Try

        If dlg.ShowDialog() <> True Then Return Nothing

        Try
            If Not String.IsNullOrEmpty(themeFolder) AndAlso
               String.Equals(Path.GetDirectoryName(dlg.FileName), themeFolder,
                             StringComparison.OrdinalIgnoreCase) Then
                Return Path.GetFileName(dlg.FileName)
            End If
        Catch
        End Try
        Return dlg.FileName
    End Function

    ' Color preview handled by ColorPickerControl

    ' --------------------------------
    '  XML SAVE
    ' --------------------------------
    Private Sub SaveXmlFile()
        If String.IsNullOrEmpty(xmlPath) Then Return
        Try
            Dim settings As New Xml.XmlWriterSettings With {
                .Indent      = True,
                .IndentChars = "  ",
                .Encoding    = New System.Text.UTF8Encoding(False)
            }
            Using writer As Xml.XmlWriter = Xml.XmlWriter.Create(xmlPath, settings)
                doc.Save(writer)
            End Using
        Catch ex As Exception
            ShowFooterError("Save failed: " & ex.Message)
        End Try
    End Sub

    ' --------------------------------
    '  XML HELPERS
    ' --------------------------------
    Private Function GetVal(xpath As String) As String
        If doc.DocumentElement Is Nothing Then Return ""
        Dim node As XmlNode = doc.SelectSingleNode(xpath)
        Return If(node IsNot Nothing, node.InnerText, "")
    End Function

    Private Sub SetVal(xpath As String, value As String)
        Dim node As XmlNode = doc.SelectSingleNode(xpath)
        If node IsNot Nothing Then node.InnerText = value
    End Sub

    Private Function NodeVal(parent As XmlNode, child As String) As String
        Dim n As XmlNode = parent.SelectSingleNode(child)
        Return If(n IsNot Nothing, n.InnerText, "")
    End Function

    Private Sub SetNodeVal(parent As XmlNode, child As String, value As String)
        Dim n As XmlNode = parent.SelectSingleNode(child)
        If n IsNot Nothing Then n.InnerText = value
    End Sub

    ' --------------------------------
    '  UI HELPERS
    ' --------------------------------
    Private Sub SetComboByTag(cmb As ComboBox, tagValue As String)
        For Each item As ComboBoxItem In cmb.Items
            If item.Tag?.ToString() = tagValue Then
                cmb.SelectedItem = item
                Return
            End If
        Next
        cmb.SelectedIndex = 0
    End Sub

    Private Function ComboTag(cmb As ComboBox) As String
        Dim item As ComboBoxItem = TryCast(cmb.SelectedItem, ComboBoxItem)
        Return If(item IsNot Nothing, item.Tag?.ToString(), "0")
    End Function

    Private Sub SetSaveStatus(msg As String)
        TxtSaveStatus.Text = msg
        TxtSaveStatus.Foreground = New SolidColorBrush(Color.FromRgb(&Hff, &H99, &H33))
    End Sub

    ' --------------------------------
    '  PREVIEW PANEL
    ' --------------------------------
    Private Sub ColorPicker_ValueChanged(sender As Object, e As EventArgs)
        UpdatePreviewPanel()
    End Sub

    Private Sub CmbDateLayout_Changed(sender As Object, e As SelectionChangedEventArgs)
        UpdatePreviewPanel()
    End Sub
    Private Sub UpdatePreviewPanel()
        Try
            ' -- Lockscreen background image --
            Dim lsPath As String = ResolveFilePath(TxtLsFilePath.Text)
            If Not String.IsNullOrEmpty(lsPath) AndAlso File.Exists(lsPath) Then
                PreviewLsImage.ImageSource = LoadBitmapSafe(lsPath)
                PreviewLsBg.Visibility = Visibility.Collapsed
            Else
                PreviewLsImage.ImageSource = Nothing
                PreviewLsBg.Visibility = Visibility.Visible
            End If

            ' -- Page background image --
            Dim pageImgPath As String = ResolveFilePath(TxtPageImage.Text)
            If Not String.IsNullOrEmpty(pageImgPath) AndAlso File.Exists(pageImgPath) Then
                PreviewPageImage.ImageSource = LoadBitmapSafe(pageImgPath)
                PreviewPageBg.Visibility = Visibility.Collapsed
            Else
                PreviewPageImage.ImageSource = Nothing
                PreviewPageBg.Visibility = Visibility.Visible
            End If

            ' -- Lockscreen notification bar color --
            ApplyPreviewColor(PreviewNotifyBarBrush,  TxtNotifyBgColor.ArgbValue)
            ApplyPreviewColor(PreviewNotifyFontBrush, TxtNotifyFontColor.ArgbValue)
            ApplyPreviewColor(PreviewClockBrush,      TxtDateColor.ArgbValue)
            ApplyPreviewColor(PreviewClockBrush2,     TxtDateColor.ArgbValue)

            ' -- Clock position --
            Dim clockStack = TryCast(PreviewClockText.Parent, StackPanel)
            If clockStack IsNot Nothing Then
                Select Case ComboTag(CmbDateLayout)
                    Case "0" ' Bottom Left
                        clockStack.VerticalAlignment   = VerticalAlignment.Bottom
                        clockStack.HorizontalAlignment = HorizontalAlignment.Left
                        clockStack.Margin = New Thickness(8, 0, 0, 8)
                    Case "1" ' Top Left
                        clockStack.VerticalAlignment   = VerticalAlignment.Top
                        clockStack.HorizontalAlignment = HorizontalAlignment.Left
                        clockStack.Margin = New Thickness(8, 22, 0, 0)
                    Case "2" ' Bottom Right
                        clockStack.VerticalAlignment   = VerticalAlignment.Bottom
                        clockStack.HorizontalAlignment = HorizontalAlignment.Right
                        clockStack.Margin = New Thickness(0, 0, 8, 8)
                End Select
            End If

            ' -- LiveArea bar colors --
            ApplyPreviewColor(PreviewBarBrush,       TxtBarColor.ArgbValue)
            ApplyPreviewColor(PreviewBarFontBrush,   TxtNoticeFontColor.ArgbValue)
            ApplyPreviewColor(PreviewIndicatorBrush, TxtIndicatorColor.ArgbValue)

            ' -- Bubble font color --
            ApplyPreviewColor(PreviewBubbleFontBrush, TxtPageFontColor.ArgbValue)
            PreviewBubbleFontHex.Text = TxtPageFontColor.ArgbValue

            ' Apply bubble tint from page font color
            Dim bubbleColor As Color = ParseArgbColor(TxtPageFontColor.ArgbValue)
            Dim bubbleTint As Color = Color.FromArgb(30, bubbleColor.R, bubbleColor.G, bubbleColor.B)
            For Each brush In {BubbleBg1, BubbleBg2, BubbleBg3, BubbleBg4,
                                BubbleBg5, BubbleBg6, BubbleBg7, BubbleBg8}
                brush.Color = bubbleTint
            Next

        Catch
            ' Silently ignore preview errors
        End Try
    End Sub

    Private Sub ApplyPreviewColor(brush As SolidColorBrush, argbHex As String)
        Try
            If Not String.IsNullOrEmpty(argbHex) AndAlso argbHex.Length = 8 Then
                Dim a As Byte = Convert.ToByte(argbHex.Substring(0, 2), 16)
                Dim r As Byte = Convert.ToByte(argbHex.Substring(2, 2), 16)
                Dim g As Byte = Convert.ToByte(argbHex.Substring(4, 2), 16)
                Dim b As Byte = Convert.ToByte(argbHex.Substring(6, 2), 16)
                brush.Color = Color.FromArgb(a, r, g, b)
            End If
        Catch
        End Try
    End Sub

    Private Function ParseArgbColor(argbHex As String) As Color
        Try
            If Not String.IsNullOrEmpty(argbHex) AndAlso argbHex.Length = 8 Then
                Dim a As Byte = Convert.ToByte(argbHex.Substring(0, 2), 16)
                Dim r As Byte = Convert.ToByte(argbHex.Substring(2, 2), 16)
                Dim g As Byte = Convert.ToByte(argbHex.Substring(4, 2), 16)
                Dim b As Byte = Convert.ToByte(argbHex.Substring(6, 2), 16)
                Return Color.FromArgb(a, r, g, b)
            End If
        Catch
        End Try
        Return Colors.Transparent
    End Function

    Private Function SanitiseName(name As String) As String
        Dim result As String = name.Trim()
        For Each c As Char In Path.GetInvalidFileNameChars()
            result = result.Replace(c, "_"c)
        Next
        Return If(String.IsNullOrEmpty(result), "NewTheme", result)
    End Function

    ' --------------------------------
    '  DEFAULT XML TEMPLATE
    '  (uses string concat to avoid
    '   interpolation quote conflicts)
    ' --------------------------------
    Private Function BuildDefaultXml(themeName As String, author As String, ver As String) As String
        Dim q As String = Chr(34)  ' double-quote character

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("<theme format-ver=" & q & "01.00" & q & " package=" & q & "0" & q & ">")
        sb.AppendLine("<InfomationProperty>")
        sb.AppendLine("  <m_contentVer>" & ver & "</m_contentVer>")
        sb.AppendLine("  <m_homePreviewFilePath>preview_page.png</m_homePreviewFilePath>")
        sb.AppendLine("  <m_packageImageFilePath>preview_thumbnail.png</m_packageImageFilePath>")
        sb.AppendLine("  <m_provider>")
        sb.AppendLine("    <m_default>" & author & "</m_default>")
        sb.AppendLine("    <m_param/>")
        sb.AppendLine("  </m_provider>")
        sb.AppendLine("  <m_startPreviewFilePath>preview_lockscreen.png</m_startPreviewFilePath>")
        sb.AppendLine("  <m_title>")
        sb.AppendLine("    <m_default>" & themeName & "</m_default>")
        sb.AppendLine("    <m_param>")
        For Each lang In {"m_da", "m_de", "m_es", "m_fi", "m_fr", "m_it",
                          "m_nl", "m_no", "m_pl", "m_pt", "m_ru", "m_sv"}
            sb.AppendLine("      <" & lang & ">" & themeName & "</" & lang & ">")
        Next
        sb.AppendLine("    </m_param>")
        sb.AppendLine("  </m_title>")
        sb.AppendLine("</InfomationProperty>")
        sb.AppendLine("<StartScreenProperty>")
        sb.AppendLine("  <m_dateColor>ffffffff</m_dateColor>")
        sb.AppendLine("  <m_dateLayout>0</m_dateLayout>")
        sb.AppendLine("  <m_filePath>lockscreen.png</m_filePath>")
        sb.AppendLine("  <m_notifyBgColor>ff000000</m_notifyBgColor>")
        sb.AppendLine("  <m_notifyBorderColor>ffcccccc</m_notifyBorderColor>")
        sb.AppendLine("  <m_notifyFontColor>ffffffff</m_notifyFontColor>")
        sb.AppendLine("</StartScreenProperty>")
        sb.AppendLine("<InfomationBarProperty>")
        sb.AppendLine("  <m_barColor>ff000000</m_barColor>")
        sb.AppendLine("  <m_indicatorColor>ffff7d00</m_indicatorColor>")
        sb.AppendLine("  <m_noticeFontColor>ffffffff</m_noticeFontColor>")
        sb.AppendLine("  <m_noticeGlowColor>ffca0000</m_noticeGlowColor>")
        sb.AppendLine("  <m_noNoticeFilePath>notices.png</m_noNoticeFilePath>")
        sb.AppendLine("  <m_newNoticeFilePath>notice.png</m_newNoticeFilePath>")
        sb.AppendLine("</InfomationBarProperty>")
        sb.AppendLine("<HomeProperty>")
        sb.AppendLine("  <m_bgParam>")
        For i As Integer = 1 To 10
            sb.AppendLine("    <BackgroundParam>")
            sb.AppendLine("      <m_thumbnailFilePath></m_thumbnailFilePath>")
            sb.AppendLine("      <m_imageFilePath></m_imageFilePath>")
            sb.AppendLine("      <m_waveType>11</m_waveType>")
            sb.AppendLine("      <m_fontColor>ffffffff</m_fontColor>")
            sb.AppendLine("      <m_fontShadow>1</m_fontShadow>")
            sb.AppendLine("    </BackgroundParam>")
        Next
        sb.AppendLine("  </m_bgParam>")
        sb.AppendLine("  <m_basePageFilePath>basePage.png</m_basePageFilePath>")
        sb.AppendLine("  <m_curPageFilePath>curPage.png</m_curPageFilePath>")
        sb.AppendLine("  <m_bgmFilePath>bgm.at9</m_bgmFilePath>")

        Dim icons As New Dictionary(Of String, String) From {
            {"m_browser",     ""},
            {"m_calendar",    ""},
            {"m_camera",      ""},
            {"m_email",       ""},
            {"m_friend",      ""},
            {"m_hostCollabo", ""},
            {"m_message",     ""},
            {"m_music",       ""},
            {"m_near",        ""},
            {"m_parental",    ""},
            {"m_party",       ""},
            {"m_ps3Link",     ""},
            {"m_ps4Link",     ""},
            {"m_power",       ""},
            {"m_settings",    ""},
            {"m_trophy",      ""},
            {"m_video",       ""}
        }

        For Each kvp In icons
            sb.AppendLine("  <" & kvp.Key & "><m_iconFilePath>" & kvp.Value & "</m_iconFilePath></" & kvp.Key & ">")
        Next

        sb.AppendLine("</HomeProperty>")
        sb.AppendLine("</theme>")
        Return sb.ToString()
    End Function

End Class
