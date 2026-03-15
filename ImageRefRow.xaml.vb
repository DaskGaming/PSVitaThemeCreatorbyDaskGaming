Imports System.Windows

Public Class ImageRefRow
    Inherits System.Windows.Controls.UserControl

    Public Shared ReadOnly FileNameProperty As DependencyProperty =
        DependencyProperty.Register("FileName", GetType(String), GetType(ImageRefRow),
            New PropertyMetadata("", AddressOf OnPropChanged))

    Public Shared ReadOnly UsedForProperty As DependencyProperty =
        DependencyProperty.Register("UsedFor", GetType(String), GetType(ImageRefRow),
            New PropertyMetadata("", AddressOf OnPropChanged))

    Public Shared ReadOnly ResolutionProperty As DependencyProperty =
        DependencyProperty.Register("Resolution", GetType(String), GetType(ImageRefRow),
            New PropertyMetadata("", AddressOf OnPropChanged))

    Public Shared ReadOnly MaxSizeProperty As DependencyProperty =
        DependencyProperty.Register("MaxSize", GetType(String), GetType(ImageRefRow),
            New PropertyMetadata("", AddressOf OnPropChanged))

    Public Shared ReadOnly NotesProperty As DependencyProperty =
        DependencyProperty.Register("Notes", GetType(String), GetType(ImageRefRow),
            New PropertyMetadata("", AddressOf OnPropChanged))

    Public Property FileName As String
        Get
            Return CStr(GetValue(FileNameProperty))
        End Get
        Set(v As String)
            SetValue(FileNameProperty, v)
        End Set
    End Property

    Public Property UsedFor As String
        Get
            Return CStr(GetValue(UsedForProperty))
        End Get
        Set(v As String)
            SetValue(UsedForProperty, v)
        End Set
    End Property

    Public Property Resolution As String
        Get
            Return CStr(GetValue(ResolutionProperty))
        End Get
        Set(v As String)
            SetValue(ResolutionProperty, v)
        End Set
    End Property

    Public Property MaxSize As String
        Get
            Return CStr(GetValue(MaxSizeProperty))
        End Get
        Set(v As String)
            SetValue(MaxSizeProperty, v)
        End Set
    End Property

    Public Property Notes As String
        Get
            Return CStr(GetValue(NotesProperty))
        End Get
        Set(v As String)
            SetValue(NotesProperty, v)
        End Set
    End Property

    Private Shared Sub OnPropChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        CType(d, ImageRefRow).UpdateLabels()
    End Sub

    Private Sub UpdateLabels()
        If TxtFileName Is Nothing Then Return
        TxtFileName.Text   = FileName
        TxtUsedFor.Text    = UsedFor
        TxtResolution.Text = Resolution
        TxtMaxSize.Text    = MaxSize
        TxtNotes.Text      = Notes
    End Sub

    Private Sub ImageRefRow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        UpdateLabels()
    End Sub

End Class
