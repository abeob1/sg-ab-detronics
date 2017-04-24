Option Strict Off
Option Explicit On
Public Class Connection
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public ocompany As New SAPbobsCOM.Company
    Public SboGuiApi As New SAPbouiCOM.SboGuiApi
    Public sConnectionString As String
    Dim oF_JournalEntry As F_JournalEntry
  
#Region "Connection"
    Public Sub New()
        MyBase.new()
        Try
            conn2()
            oF_JournalEntry = New F_JournalEntry(ocompany, SBO_Application)
            SBO_Application.MessageBox("Welcome To JE...")
            SBO_Application.StatusBar.SetText("Add-on Loaded!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:New" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub conn1()
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String
        Try
            sconn = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sconn)
            SBO_Application = SboGuiApi.GetApplication
            SboGuiApi = Nothing
            scook = ocompany.GetContextCookie
            str = SBO_Application.Company.GetConnectionContext(scook)
            ret = ocompany.SetSboLoginContext(str)
            ocompany.Connect()
            ocompany.GetLastError(ret, str)
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:conn1" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub conn2()
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String
        Try
            sconn = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sconn)
            SBO_Application = SboGuiApi.GetApplication
            SboGuiApi = Nothing
            scook = ocompany.GetContextCookie
            str = SBO_Application.Company.GetConnectionContext(scook)
            ret = ocompany.SetSboLoginContext(str)
            ocompany.Connect()
            ocompany.GetLastError(ret, str)
            If ret <> 0 Then
                SBO_Application.MessageBox("SAP Connection Failed :" & str)
                Functions.WriteLog("Class:Connection" + " Function:conn2" + " Error Message:" + str)
            End If
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:conn2" + " Error Message:" + ex.ToString)
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
#End Region
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        Try
            If (EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_FontChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged) Then
                SBO_Application.StatusBar.SetText("Shuting Down addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Windows.Forms.Application.Exit()
            End If
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:SBO_Application_AppEvent" + " Error Message:" + ex.ToString)
        End Try
    End Sub
End Class
