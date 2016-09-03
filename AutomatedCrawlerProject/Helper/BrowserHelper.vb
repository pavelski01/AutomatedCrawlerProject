Imports System.Runtime.InteropServices
Imports System.Threading
Imports mshtml
Imports SHDocVw

Namespace Helper
    Public Module BrowserHelper
        Public Function GetHtmlDocument(pageTitle As String) As HTMLDocument
            For index = 0 To 50
                For Each retIe As ShellBrowserWindow In
                From ieSw As ShellBrowserWindow In New ShellWindows
                Where
                    ieSw.Document IsNot Nothing AndAlso
                    TypeOf ieSw.Document Is HTMLDocument AndAlso
                    CType(ieSw.Document, HTMLDocument).title.Trim().StartsWith(pageTitle)
                    Return retIe.Document
                Next
                Thread.Sleep(100)
            Next
            Return Nothing
        End Function
        Public Function GetElementById(document As HTMLDocument, id As String) As IHTMLElement
            WaitForDocumentReadyState(document)
            While document.getElementById(id) Is Nothing
                Thread.Sleep(100)
            End While
            Return document.getElementById(id)
        End Function
        Public Function GetHtmlElementsByClass(document As HTMLDocument, tag As String, clazzName As String) As List(Of IHTMLElement)
            While document.getElementsByTagName(tag) Is Nothing
                Thread.Sleep(100)
            End While
            Dim unfilteredElementCollection = document.getElementsByTagName(tag)
            Return (
                From elem As IHTMLElement In unfilteredElementCollection
                Where elem.getAttribute("class") IsNot Nothing AndAlso elem.getAttribute("class") = clazzName
            ).ToList()
        End Function
        Public Function CreateMaximizedBrowser() As InternetExplorer
            Dim ie As New InternetExplorer With {.Visible = True}
            While ie Is Nothing
                Thread.Sleep(100)
            End While
            MaximizeForeground(ie.HWND)
            Return ie
        End Function
        Public Function SafeNavigate(ie As InternetExplorer, url As String) As Boolean
            ie.Navigate(url)
            WaitForInternetExplorerReadyState(ie)
            While ie.Document Is Nothing
                Thread.Sleep(100)
            End While
            WaitForDocumentReadyState(ie.Document)
            Return True
        End Function
        Public Sub WaitForInternetExplorerReadyState(ie As InternetExplorer)
            Dim index = 0
            Try
                index += 1
                While ie.ReadyState <> 4 AndAlso ie.ReadyState <> 3 AndAlso index < 50
                    Thread.Sleep(100)
                End While
            Catch ex As Exception
                WaitForInternetExplorerReadyState(ie)
            End Try
        End Sub
        Public Sub WaitForDocumentReadyState(document As HTMLDocument)
            While Not String.Equals(document.readyState, "complete")
                Thread.Sleep(100)
            End While
        End Sub
        Public Sub WaitForDocumentReadyState(ie As InternetExplorer)
            While ie.Document Is Nothing
                Thread.Sleep(100)
            End While
            WaitForDocumentReadyState(ie.Document)
        End Sub
        Public Function TraverseToNode(elem As IHTMLElement, childrenList As Integer()) As IHTMLElement
            Dim endRoad = elem
            For Each nr In childrenList
                If nr >= 0 Then
                    endRoad = endRoad.children(nr)
                Else
                    endRoad = endRoad.parentElement
                End If
            Next
            Return endRoad
        End Function
        Public Function TraverseToNode(elem As IHTMLElement, childNr As Integer) As IHTMLElement
            Dim childrenList = New Integer() {childNr}
            Return TraverseToNode(elem, childrenList)
        End Function
        <DllImport("user32.dll")>
        Private Sub SetForegroundWindow(hwnd As IntPtr)
        End Sub
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Private Sub ShowWindow(hwnd As IntPtr, nCmdShow As Integer)
        End Sub
        Private Sub MaximizeForeground(hwnd As IntPtr)
            SetForegroundWindow(hwnd)
            ShowWindow(hwnd, 3)
        End Sub
    End Module
End Namespace
