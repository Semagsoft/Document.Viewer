Imports System.IO
Class MainWindow
    Private web As New WebBrowser, Speech As New Speech.Synthesis.SpeechSynthesizer

    Private Sub ReadDocx(path__1 As String)
        Using stream = File.Open(path__1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim flowDocumentConverter = New DocxReaderApplication.DocxToFlowDocumentConverter(stream)
            flowDocumentConverter.Read()
            DocumentViewer.Document = flowDocumentConverter.Document
        End Using
    End Sub

#Region "Themes"

    Private Enum Theme
        Office2010
        Office2013
    End Enum

    Private currentTheme As System.Nullable(Of Theme)

    Private Sub SetRibbonTheme_Office2013(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2013, "pack://application:,,,/Fluent;component/Themes/Office2013/Generic.xaml")
    End Sub

    Private Sub SetRibbonTheme__Office2010Silver(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2010, "pack://application:,,,/Fluent;component/Themes/Office2010/Silver.xaml")
    End Sub

    Private Sub SetRibbonTheme_Office2010Black(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2010, "pack://application:,,,/Fluent;component/Themes/Office2010/Black.xaml")
    End Sub

    Private Sub SetRibbonTheme_Office2010Blue(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2010, "pack://application:,,,/Fluent;component/Themes/Office2010/Blue.xaml")
    End Sub

    Private Sub ChangeTheme(theme_1 As Theme, color As String)
        Me.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, DirectCast(Function()

                                                                                                              Dim owner = Window.GetWindow(Me)
                                                                                                              If owner IsNot Nothing Then
                                                                                                                  owner.Resources.BeginInit()

                                                                                                                  If owner.Resources.MergedDictionaries.Count > 0 Then
                                                                                                                      owner.Resources.MergedDictionaries.RemoveAt(0)
                                                                                                                  End If

                                                                                                                  If String.IsNullOrEmpty(color) = False Then
                                                                                                                      owner.Resources.MergedDictionaries.Add(New ResourceDictionary() With { _
                                                                                                                          .Source = New Uri(color) _
                                                                                                                      })
                                                                                                                  End If

                                                                                                                  owner.Resources.EndInit()
                                                                                                              End If

                                                                                                              If Me.currentTheme <> theme_1 Then
                                                                                                                  Application.Current.Resources.BeginInit()
                                                                                                                  Select Case theme_1
                                                                                                                      Case Theme.Office2010
                                                                                                                          Application.Current.Resources.MergedDictionaries.Add(New ResourceDictionary() With { _
                                                                                                                              .Source = New Uri("pack://application:,,,/Fluent;component/Themes/Generic.xaml") _
                                                                                                                          })
                                                                                                                          Application.Current.Resources.MergedDictionaries.RemoveAt(0)
                                                                                                                          Exit Select
                                                                                                                      Case Theme.Office2013
                                                                                                                          Application.Current.Resources.MergedDictionaries.Add(New ResourceDictionary() With { _
                                                                                                                              .Source = New Uri("pack://application:,,,/Fluent;component/Themes/Office2013/Generic.xaml") _
                                                                                                                          })
                                                                                                                          Application.Current.Resources.MergedDictionaries.RemoveAt(0)
                                                                                                                          Exit Select
                                                                                                                  End Select

                                                                                                                  Me.currentTheme = theme_1
                                                                                                                  Application.Current.Resources.EndInit()

                                                                                                                  If owner IsNot Nothing Then
                                                                                                                      owner.Style = Nothing
                                                                                                                      owner.Style = TryCast(owner.FindResource("RibbonWindowStyle"), Style)
                                                                                                                      owner.Style = Nothing
                                                                                                                  End If
                                                                                                              End If

                                                                                                          End Function, System.Threading.ThreadStart))
    End Sub

#End Region

#Region "MainWindow"

    Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        If Me.WindowState = Windows.WindowState.Maximized Then
            My.Settings.MainWindow_IsMax = True
        Else
            My.Settings.MainWindow_IsMax = False
            My.Settings.MainWindow_Size = New Size(Me.Width, Me.Height)
        End If
        My.Settings.Save()
    End Sub

    Private Sub MainWindow_KeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs) Handles Me.KeyDown
        If Keyboard.IsKeyDown(Key.LeftCtrl) OrElse Keyboard.IsKeyDown(Key.Right) Then
            If e.Key = Key.O Then
                OpenButton_Click(Nothing, Nothing)
            ElseIf e.Key = Key.P Then
                PrintButton_Click(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Settings.MainWindow_IsMax Then
            Me.WindowState = Windows.WindowState.Maximized
        Else
            Me.Width = My.Settings.MainWindow_Size.Width
            Me.Height = My.Settings.MainWindow_Size.Height
        End If
        If My.Settings.Options_Theme = 1 Then
            SetRibbonTheme_Office2010Blue(Nothing, Nothing)
        ElseIf My.Settings.Options_Theme = 2 Then
            SetRibbonTheme__Office2010Silver(Nothing, Nothing)
        ElseIf My.Settings.Options_Theme = 3 Then
            SetRibbonTheme_Office2010Black(Nothing, Nothing)
        End If
        web.HorizontalAlignment = HorizontalAlignment.Stretch
        web.VerticalAlignment = VerticalAlignment.Stretch
        ContentBox.Children.Add(web)
        XpsViewer.HorizontalAlignment = HorizontalAlignment.Stretch
        XpsViewer.VerticalAlignment = VerticalAlignment.Stretch
        If My.Application.StartUpFileNames.Count > 0 Then
            Dim f As New IO.FileInfo(My.Application.StartUpFileNames.Item(0)), fs As IO.FileStream = IO.File.OpenRead(f.FullName)
            Dim tr As New TextRange(DocumentViewer.Document.ContentStart, DocumentViewer.Document.ContentEnd)
            If f.Extension = ".xaml" Then
                web.Visibility = Visibility.Hidden
                XpsViewer.Visibility = Visibility.Hidden
                DocumentViewer.Visibility = Visibility.Visible
                Dim xamlFile As New IO.FileStream(f.FullName, IO.FileMode.Open, IO.FileAccess.Read)
                Dim content As FlowDocument = TryCast(Markup.XamlReader.Load(xamlFile), FlowDocument)
                DocumentViewer.Document = content
                xamlFile.Close()
            ElseIf f.Extension = ".docx" Then
                web.Visibility = Visibility.Hidden
                XpsViewer.Visibility = Visibility.Hidden
                DocumentViewer.Visibility = Visibility.Visible
                ReadDocx(f.FullName)
            ElseIf f.Extension = ".xps" Then
                DocumentViewer.Visibility = Visibility.Hidden
                web.Visibility = Visibility.Hidden
                XpsViewer.Visibility = Visibility.Visible
                Dim d As Xps.Packaging.XpsDocument = New Xps.Packaging.XpsDocument(f.FullName, System.IO.FileAccess.Read)
                XpsViewer.Document = d.GetFixedDocumentSequence()
            ElseIf f.Extension = ".html" Or f.Extension = ".htm" Then
                web.Visibility = Visibility.Visible
                DocumentViewer.Visibility = Visibility.Hidden
                XpsViewer.Visibility = Visibility.Hidden
                web.Source = New Uri(f.FullName)
            ElseIf f.Extension = ".rtf" Then
                DocumentViewer.Visibility = Visibility.Visible
                XpsViewer.Visibility = Visibility.Hidden
                web.Visibility = Visibility.Hidden
                Dim fileStream As IO.FileStream = IO.File.Open(f.FullName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                Dim flowDocument As New FlowDocument(), textRange As New TextRange(flowDocument.ContentStart, flowDocument.ContentEnd)
                textRange.Load(fs, DataFormats.Rtf)
                DocumentViewer.Document = flowDocument
            Else
                DocumentViewer.Visibility = Visibility.Visible
                XpsViewer.Visibility = Visibility.Hidden
                web.Visibility = Visibility.Hidden
                Dim fileStream As IO.FileStream = IO.File.Open(f.FullName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                Dim flowDocument As New FlowDocument(), textRange As New TextRange(flowDocument.ContentStart, flowDocument.ContentEnd)
                textRange.Load(fs, DataFormats.Text)
                DocumentViewer.Document = flowDocument
            End If
            fs.Close()
            If Not My.Settings.Options_RecentDocuments.Contains(f.FullName) AndAlso My.Settings.Options_RecentDocuments.Count < 13 Then
                My.Settings.Options_RecentDocuments.Add(f.FullName)
            End If
            ExportMenuItem.IsEnabled = True
            PrintMenuItem.IsEnabled = True
            PrintButton.IsEnabled = True
            ClipboardGroup.IsEnabled = True
            ZoomGroup.IsEnabled = True
            SpeechMenuItem.IsEnabled = True
        End If
        If My.Settings.Options_RecentDocuments.Count > 0 Then
            Dim grid As New StackPanel
            For Each s As String In My.Settings.Options_RecentDocuments
                If My.Computer.FileSystem.FileExists(s) Then
                    Dim i2 As New Fluent.Button, f As New IO.FileInfo(s)
                    Dim ContMenu As New ContextMenu
                    Dim removemenuitem As New MenuItem
                    removemenuitem.Header = "Remove"
                    removemenuitem.ToolTip = f.FullName
                    ContMenu.Items.Add(removemenuitem)
                    i2.ContextMenu = ContMenu
                    'Dim img As New Image
                    'If f.Extension.ToLower = ".xaml" Then
                    '    img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/xaml16.png"))
                    'ElseIf f.Extension.ToLower = ".rtf" Then
                    '    img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/rtf16.png"))
                    'ElseIf f.Extension.ToLower = ".html" OrElse f.Extension.ToLower = ".htm" Then
                    '    img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/html16.png"))
                    'Else
                    '    img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/txt16.png"))
                    'End If
                    'i2.Icon = img
                    i2.Foreground = Brushes.Black
                    i2.Header = f.Name
                    i2.Tag = s
                    RecentDocumentsList.Children.Add(i2)
                    AddHandler (i2.Click), New RoutedEventHandler(AddressOf recentfile_click)
                    AddHandler removemenuitem.Click, New RoutedEventHandler(AddressOf recentfileremove_click)
                End If
            Next
        Else
            'FileMenuItem.Items.Remove(RecentFilesTabItem)
        End If
    End Sub

#End Region

#Region "MainMenu"

#Region "File"

    Private Sub recentfile_click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
        Dim i As Fluent.Button = e.Source, f As New IO.FileInfo(i.Tag), fs As IO.FileStream = IO.File.OpenRead(f.FullName)
        If f.Extension = ".xaml" Then
            web.Visibility = Visibility.Hidden
            XpsViewer.Visibility = Visibility.Hidden
            DocumentViewer.Visibility = Visibility.Visible
            Dim xamlFile As New IO.FileStream(f.FullName, IO.FileMode.Open, IO.FileAccess.Read), content As FlowDocument = TryCast(Markup.XamlReader.Load(xamlFile), FlowDocument)
            DocumentViewer.Document = content
            xamlFile.Close()
        ElseIf f.Extension = ".docx" Then
            web.Visibility = Visibility.Hidden
            XpsViewer.Visibility = Visibility.Hidden
            DocumentViewer.Visibility = Visibility.Visible
            ReadDocx(f.FullName)
        ElseIf f.Extension = ".xps" Then
            DocumentViewer.Visibility = Visibility.Hidden
            web.Visibility = Visibility.Hidden
            XpsViewer.Visibility = Visibility.Visible
            Dim d As Xps.Packaging.XpsDocument = New Xps.Packaging.XpsDocument(f.FullName, System.IO.FileAccess.Read)
            XpsViewer.Document = d.GetFixedDocumentSequence()
        ElseIf f.Extension = ".html" Or f.Extension = ".htm" Then
            web.Visibility = Visibility.Visible
            DocumentViewer.Visibility = Visibility.Hidden
            XpsViewer.Visibility = Visibility.Hidden
            web.Source = New Uri(f.FullName)
        ElseIf f.Extension = ".rtf" Then
            DocumentViewer.Visibility = Visibility.Visible
            web.Visibility = Visibility.Hidden
            XpsViewer.Visibility = Windows.Visibility.Hidden
            Dim fileStream As IO.FileStream = IO.File.Open(f.FullName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
            Dim flowDocument As New FlowDocument(), textRange As New TextRange(flowDocument.ContentStart, flowDocument.ContentEnd)
            textRange.Load(fileStream, DataFormats.Rtf)
            DocumentViewer.Document = flowDocument
        Else
            DocumentViewer.Visibility = Visibility.Visible
            web.Visibility = Visibility.Hidden
            XpsViewer.Visibility = Visibility.Hidden
            Dim fileStream As IO.FileStream = IO.File.Open(f.FullName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
            Dim flowDocument As New FlowDocument(), textRange As New TextRange(flowDocument.ContentStart, flowDocument.ContentEnd)
            textRange.Load(fileStream, DataFormats.Text)
            DocumentViewer.Document = flowDocument
        End If
        fs.Close()
        ExportMenuItem.IsEnabled = True
        PrintMenuItem.IsEnabled = True
        PrintButton.IsEnabled = True
        ClipboardGroup.IsEnabled = True
        ZoomGroup.IsEnabled = True
        SpeechMenuItem.IsEnabled = True
    End Sub

    Private Sub recentfileremove_click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
        Dim i As MenuItem = e.Source
        Dim itemtoremove As Fluent.Button = Nothing
        For Each recentdoc As Fluent.Button In RecentDocumentsList.Children
            If recentdoc.Tag = i.ToolTip Then
                itemtoremove = recentdoc
            End If
        Next
        Dim stringtoremove As String = Nothing
        For Each s As String In My.Settings.Options_RecentDocuments
            If s = i.ToolTip Then
                stringtoremove = s
            End If
        Next
        My.Settings.Options_RecentDocuments.Remove(stringtoremove)
        RecentDocumentsList.Children.Remove(itemtoremove)
    End Sub

#Region "Open"

    Private Sub OpenButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OpenMenuItem.Click, OpenButton.Click
        Try
            Dim open As New Microsoft.Win32.OpenFileDialog
            open.Filter = "Supported Documents(*.xaml;*.docx;*.xps;*.html;*.htm;*.rtf;*.txt)|*.xaml;*.docx;*.xps;*.html;*.htm;*.rtf;*.txt|FlowDocments(*.xaml)|*.xaml|DocX Document(*.docx)|*.docx|XPS Documents(*.xps)|*.xps|HTML Documents(*.html;*.htm)|*.html;*.htm|Rich Text Documents(*.rtf)|*.rtf|Text Documents(*.txt)|*.txt|All Files(*.*)|*.*"
            If open.ShowDialog Then
                Dim f As New IO.FileInfo(open.FileName), fs As IO.FileStream = IO.File.OpenRead(f.FullName)
                If f.Extension = ".xaml" Then
                    web.Visibility = Visibility.Hidden
                    XpsViewer.Visibility = Visibility.Hidden
                    DocumentViewer.Visibility = Visibility.Visible
                    Dim xamlFile As New IO.FileStream(f.FullName, IO.FileMode.Open, IO.FileAccess.Read), content As FlowDocument = TryCast(Markup.XamlReader.Load(xamlFile), FlowDocument)
                    DocumentViewer.Document = content
                    xamlFile.Close()
                ElseIf f.Extension = ".docx" Then
                    web.Visibility = Visibility.Hidden
                    XpsViewer.Visibility = Visibility.Hidden
                    DocumentViewer.Visibility = Visibility.Visible
                    ReadDocx(f.FullName)
                ElseIf f.Extension = ".xps" Then
                    DocumentViewer.Visibility = Visibility.Hidden
                    web.Visibility = Visibility.Hidden
                    XpsViewer.Visibility = Visibility.Visible
                    Dim d As Xps.Packaging.XpsDocument = New Xps.Packaging.XpsDocument(open.FileName, System.IO.FileAccess.Read)
                    XpsViewer.Document = d.GetFixedDocumentSequence()
                ElseIf f.Extension = ".html" Or f.Extension = ".htm" Then
                    web.Visibility = Visibility.Visible
                    DocumentViewer.Visibility = Visibility.Hidden
                    XpsViewer.Visibility = Visibility.Hidden
                    web.Source = New Uri(open.FileName)
                ElseIf f.Extension = ".rtf" Then
                    DocumentViewer.Visibility = Visibility.Visible
                    web.Visibility = Visibility.Hidden
                    XpsViewer.Visibility = Windows.Visibility.Hidden
                    Dim fileStream As IO.FileStream = IO.File.Open(f.FullName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                    Dim flowDocument As New FlowDocument(), textRange As New TextRange(flowDocument.ContentStart, flowDocument.ContentEnd)
                    textRange.Load(fileStream, DataFormats.Rtf)
                    DocumentViewer.Document = flowDocument
                Else
                    DocumentViewer.Visibility = Visibility.Visible
                    web.Visibility = Visibility.Hidden
                    XpsViewer.Visibility = Visibility.Hidden
                    Dim fileStream As IO.FileStream = IO.File.Open(f.FullName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                    Dim flowDocument As New FlowDocument(), textRange As New TextRange(flowDocument.ContentStart, flowDocument.ContentEnd)
                    textRange.Load(fileStream, DataFormats.Text)
                    DocumentViewer.Document = flowDocument
                End If
                fs.Close()
                If Not My.Settings.Options_RecentDocuments.Contains(f.FullName) AndAlso My.Settings.Options_RecentDocuments.Count < 13 Then
                    My.Settings.Options_RecentDocuments.Add(f.FullName)
                End If
                ExportMenuItem.IsEnabled = True
                PrintMenuItem.IsEnabled = True
                PrintButton.IsEnabled = True
                ClipboardGroup.IsEnabled = True
                ZoomGroup.IsEnabled = True
                SpeechMenuItem.IsEnabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region

#Region "Export"

    Private Sub ExportTextFileMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ExportTextFileMenuItem.Click
        Dim export As New Microsoft.Win32.SaveFileDialog
        export.Filter = "Text File(*.txt)|*.txt|All Files(*.*)|*.*"
        If export.ShowDialog Then
            If DocumentViewer.Visibility = Windows.Visibility.Visible Then
                Dim tr As New TextRange(DocumentViewer.Document.ContentStart, DocumentViewer.Document.ContentEnd)
                My.Computer.FileSystem.WriteAllText(export.FileName, tr.Text, False)
            ElseIf XpsViewer.Visibility = Windows.Visibility.Visible Then
                'TODO: fix for xps documents

            ElseIf web.Visibility = Windows.Visibility.Visible Then
                'TODO: fix for html
                'My.Computer.FileSystem.WriteAllText(export.FileName, web.PageContents, False)
            End If
        End If
    End Sub

    Private Sub ExportImageMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ExportImageMenuItem.Click
        Dim export As New Microsoft.Win32.SaveFileDialog
        export.Filter = "Png Image|*.png|All Files|*.*"
        If export.ShowDialog Then
            If DocumentViewer.Visibility = Windows.Visibility.Visible Then
                'TODO: fix for document viewer

            ElseIf XpsViewer.Visibility = Windows.Visibility.Visible Then
                'TODO: fix for xps

            ElseIf web.Visibility = Windows.Visibility.Visible Then
                'TODO: fix for html
                'web.SaveToPNG(export.FileName)
            End If
        End If
    End Sub

    Private Sub ExportAudioMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ExportAudioMenuItem.Click
        Dim export As New Microsoft.Win32.SaveFileDialog

        If export.ShowDialog Then
            'TODO:

        End If
    End Sub

#End Region

    Private Sub PrintButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles PrintMenuItem.Click, PrintButton.Click
        If DocumentViewer.IsVisible Then
            Dim pd As New PrintDialog
            pd.PrintVisual(DocumentViewer, "test")
        ElseIf XpsViewer.IsVisible Then
            XpsViewer.Print()
        ElseIf web.IsVisible Then
            Dim pd As New PrintDialog
            pd.PrintVisual(web, "test")
        End If
    End Sub

    Private Sub ExitMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ExitMenuItem.Click
        Close()
    End Sub

#End Region

#Region "Edit"

    Private Sub CopyMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles CopyMenuItem.Click
        If DocumentViewer.IsVisible Then
            Clipboard.SetText(DocumentViewer.Selection.Text)
        ElseIf XpsViewer.IsVisible Then
            ApplicationCommands.Copy.Execute(Nothing, XpsViewer)
        ElseIf web.IsVisible Then
            'TODO: fix for html
            'web.Copy()
        End If
    End Sub

    Private Sub SelectAllMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SelectAllMenuItem.Click
        If DocumentViewer.IsVisible Then
            DocumentViewer.Selection.Select(DocumentViewer.Document.ContentStart, DocumentViewer.Document.ContentEnd)
        ElseIf XpsViewer.IsVisible Then
            ApplicationCommands.SelectAll.Execute(Nothing, XpsViewer)
        ElseIf web.IsVisible Then
            'TODO: fix for html
            'web.SelectAll()
        End If
    End Sub

#End Region

#Region "View"

    Private Sub ZoomInMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles ZoomInMenuItem.Click
        If DocumentViewer.IsVisible Then
            If DocumentViewer.CanIncreaseZoom Then
                DocumentViewer.IncreaseZoom()
            End If
        ElseIf XpsViewer.IsVisible Then
            If XpsViewer.CanIncreaseZoom Then
                XpsViewer.IncreaseZoom()
            End If
        Else 'TODO: fix for html

        End If
        UpdateZoom()
    End Sub

    Private Sub ZoomOutMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles ZoomOutMenuItem.Click
        If DocumentViewer.IsVisible Then
            If DocumentViewer.CanDecreaseZoom Then
                DocumentViewer.DecreaseZoom()
            End If
        ElseIf XpsViewer.IsVisible Then
            If XpsViewer.CanDecreaseZoom Then
                XpsViewer.DecreaseZoom()
            End If
        Else 'TODO: fix for html

        End If
        UpdateZoom()
    End Sub

    Private Sub ResetZoomMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles ResetZoomMenuItem.Click
        If DocumentViewer.IsVisible Then
            DocumentViewer.Zoom = 100
        ElseIf XpsViewer.IsVisible Then
            XpsViewer.Zoom = 100
        Else 'TODO: fix for html

        End If
        UpdateZoom()
    End Sub

    Private Sub UpdateZoom()
        If DocumentViewer.IsVisible Then
            If DocumentViewer.CanIncreaseZoom Then
                ZoomInMenuItem.IsEnabled = True
            Else
                ZoomInMenuItem.IsEnabled = False
            End If
            If DocumentViewer.CanDecreaseZoom Then
                ZoomOutMenuItem.IsEnabled = True
            Else
                ZoomOutMenuItem.IsEnabled = False
            End If
        ElseIf XpsViewer.IsVisible Then
            If XpsViewer.CanIncreaseZoom Then
                ZoomInMenuItem.IsEnabled = True
            Else
                ZoomInMenuItem.IsEnabled = False
            End If
            If XpsViewer.CanDecreaseZoom Then
                ZoomOutMenuItem.IsEnabled = True
            Else
                ZoomOutMenuItem.IsEnabled = False
            End If
        Else 'TODO: fix for html

        End If
    End Sub

#End Region

#Region "Tools"

#Region "Speech"

    Private Sub SpeechMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SpeechMenuItem.Click
        If DocumentViewer.IsVisible Then
            If Speech.State = System.Speech.Synthesis.SynthesizerState.Speaking Then
                Speech.Pause()
            Else
                Try
                    Speech.SelectVoice(Speech.GetInstalledVoices.Item(0).VoiceInfo.Name)
                    Speech.Resume()
                    Speech.SpeakAsync(DocumentViewer.Selection.Text)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Text to Speech Error")
                End Try
            End If
        ElseIf web.IsVisible Then
            'TODO: fix Speech for xps documents

        ElseIf web.IsVisible Then
            'TODO: fix for html
            'If Speech.State = SynthesizerState.Speaking Then
            '    Speech.Pause()
            'Else
            '    Try
            '        Speech.SelectVoice(Speech.GetInstalledVoices.Item(0).VoiceInfo.Name)
            '        Speech.Resume()
            '        Speech.SpeakAsync(web.Selection.Text)
            '    Catch ex As Exception
            '        MessageBox.Show(ex.Message, "Text to Speech Error")
            '    End Try
            'End If
        End If
    End Sub

#End Region

    Private Sub OptionsMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles OptionsMenuItem.Click
        Dim o As New OptionsDialog
        o.Owner = Me
        o.ShowDialog()
    End Sub

#End Region

#Region "Help"

    Private Sub OnlineHelpMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles OnlineHelpMenuItem.Click
        Process.Start("http://documentviewer.codeplex.com")
    End Sub

    Private Sub WebsiteMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles WebsiteMenuItem.Click
        Process.Start("http://semagsoft.com")
    End Sub

    Private Sub DonateMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles DonateMenuItem.Click
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=K4QJBR4UJ3W5E&lc=US&item_name=Semagsoft&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donateCC_LG%2egif%3aNonHosted")
    End Sub

    Private Sub AboutMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles AboutMenuItem.Click
        Dim a As New AboutDialog
        a.Owner = Me
        a.ShowDialog()
    End Sub

#End Region

#End Region

End Class