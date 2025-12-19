Partial Public Class AboutDialog

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Close()
    End Sub

    Private Sub AboutDialog_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            GlassHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
        End If

        VersionLabel.Content = "Version: " + My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString
        CopyLabel.Content = My.Application.Info.Copyright.ToString + " By Semagsoft"
        TextBox1.Text = My.Resources.License
    End Sub
End Class