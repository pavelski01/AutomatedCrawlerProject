'Imports AutomatedCrawlerProject.Helper.BrowserHelper
Imports AutomatedCrawlerProject.Helper.MainframeHelper
Partial Public Class MainWindow
    Inherits Window
    Sub New()
        InitializeComponent()
    End Sub
    Private Sub ButtonBase_OnClick(sender As Object, e As RoutedEventArgs)
        Dim conn = ConnectionHandler()
        Dim ps = GetPresentationSpace(conn)
        WaitForText(ps, "Mainframe Operating System")
        SetCursorPosition(ps, 21, 22)
        SendKeys(ps, "TSO")
        SendKeys(ps, "[enter]")
        WaitForText(ps, "ENTER USERID")
        SetCursorPosition(ps, 2, 1)
        SendKeys(ps, "PAVEL1")
        SendKeys(ps, "[enter]")
        WaitForText(ps, "TSO/E LOGON")
        SetCursorPosition(ps, 21, 11)
        SendKeys(ps, "S")
        SendKeys(ps, "[enter]")
    End Sub
End Class
