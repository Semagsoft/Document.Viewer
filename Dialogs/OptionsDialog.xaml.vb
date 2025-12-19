Public Class OptionsDialog

    Private Sub OptionsDialog_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ThemeCombeBox.SelectedIndex = My.Settings.Options_Theme
    End Sub

    Private Sub OKButton_Click(sender As Object, e As RoutedEventArgs) Handles OKButton.Click
        My.Settings.Options_Theme = ThemeCombeBox.SelectedIndex
        Close()
    End Sub
End Class