Imports System.Threading
Imports AutConnListTypeLibrary
Imports AutConnMgrTypeLibrary
Imports AutPSTypeLibrary

Namespace Helper
    Public Module MainframeHelper
        Public Function ConnectionHandler() As IAutConnInfo
            Dim mgr As IAutConnMgr = CreateObject("PCOMM.autECLConnMgr")
            Dim autEclConnList As IAutConnList = mgr.autECLConnList
            Dim connObj As IAutConnInfo
            Dim hand As Long
            autEclConnList.Refresh()
            If autEclConnList(1).Ready Then
                hand = autEclConnList(1).Handle
                connObj = autEclConnList.FindConnectionByHandle(hand)
                Return connObj
            Else
                Return Nothing
            End If
        End Function
        Public Function GetPresentationSpace(connObj As IAutConnInfo) As IAutPS
            Return GetPresentationSpace(Convert.ToInt64(connObj.Handle))
        End Function
        Public Function GetPresentationSpace(hand As Long) As IAutPS
            Dim autEclPsObj As IAutPS = CreateObject("PCOMM.autECLPS")
            autEclPsObj.SetConnectionByHandle(hand)
            Return autEclPsObj
        End Function
        Public Sub SetCursorPosition(ps As IAutPS, row As Integer, column As Integer)
            ps.SetCursorPos(row, column)
        End Sub
        Public Sub SendKeys(ps As IAutPS, value As String)
            ps.SendKeys(value)
        End Sub
        Public Sub SendKeys(ps As IAutPS, value As String, row As Integer, col As Integer)
            ps.SendKeys(value, row, col)
        End Sub
        Public Function WaitForText(ps As IAutPS, value As String) As Boolean
            While Not ps.SearchText(value)
                Thread.Sleep(100)
            End While
            Return True
        End Function
    End Module
End Namespace