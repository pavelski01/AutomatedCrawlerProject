Imports AutomatedCrawlerProject.Helper.BrowserHelper

Partial Public Class MainWindow
    Inherits Window
    Sub New()
        InitializeComponent()
    End Sub
    Private Sub ButtonBase_OnClick(sender As Object, e As RoutedEventArgs)
        Dim ie = CreateMaximizedBrowser()
        SafeNavigate(ie, "http://niezalezna.pl")
        Dim a = GetHtmlDocument("Wiadomości, informacje, publicystyka, opinie | niezalezna.pl")
    End Sub
End Class
