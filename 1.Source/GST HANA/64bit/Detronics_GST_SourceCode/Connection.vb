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
            'SBO_Application.MessageBox("Welcome To JE...")
            SBO_Application.StatusBar.SetText("Add-on Loaded!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            Functions.WriteLog("Class:Connection" + " Function:New" + " Error Message:" + ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub conn2()
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String
        Try
            sconn = ""
            Try
                If Environment.GetCommandLineArgs.Length > 1 Then
                    sconn = Environment.GetCommandLineArgs.GetValue(1)
                Else
                    sconn = Environment.GetCommandLineArgs.GetValue(0)
                End If

            Catch ex As Exception

            End Try
            'sconn = CStr(Environment.GetCommandLineArgs.GetValue(1))
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
    ' Private Sub JournalEntry(ByVal DocEntry As String)
    '    Try
    '        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRecordSet.DoQuery("SELECT [TransId],U_AP_EXCH_RATE,U_AP_TAX_AMT FROM OPCH WHERE DocEntry='" & DocEntry & "'")
    '        oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim ExRate As Double = 0
    '        Dim TaxAmt As Double = 0
    '        ExRate = oRecordSet.Fields.Item("U_AP_EXCH_RATE").Value
    '        TaxAmt = oRecordSet.Fields.Item("U_AP_TAX_AMT").Value
    '        oRecordSet1.DoQuery("SELECT T0.[RefDate], T0.[DueDate], T0.[TaxDate], T0.[Memo] , 'PU' 'Indicator' ,T0.Ref1,T0.[TransId], T1.[Line_ID],CASE WHEN T1.[Account]= T1.[ShortName]  THEN 'N' ELSE 'Y' END as 'SN',T1.[ShortName] ,CASE WHEN T1.[SYSCred] <> 0 THEN 'Credit' ELSE 'Debit' END AS 'SIDE',(CASE WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCCredit]=0  THEN T1.[FCDebit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Credit]=0  THEN T1.[Debit] WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCDebit]=0  THEN T1.[FCCredit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Debit]=0  THEN T1.[Credit] END)  * " & ExRate & "- (CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END)As 'LOCALAMOUNT',CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END AS 'SYSTEMAMOUNT',((T1.[BaseSum]*" & ExRate & ")- T1.[SYSBaseSum])'BASEAMOUNT',ISNULL(T1.VatGroup,'') 'VatGroup' FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId WHERE T1.[TransId]='" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.[TransId], T1.[Line_ID] ASC")
    '        Dim oJV As SAPbobsCOM.JournalEntries
    '        oJV = ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    '        Dim i As Integer = 0
    '        Dim TotDebit As Double = 0
    '        Dim TotCredit As Double = 0
    '        If oRecordSet1.RecordCount = 0 Then
    '            Exit Sub
    '        End If
    '        oJV.ReferenceDate = oRecordSet1.Fields.Item("RefDate").Value
    '        oJV.DueDate = oRecordSet1.Fields.Item("DueDate").Value
    '        oJV.TaxDate = oRecordSet1.Fields.Item("TaxDate").Value
    '        oJV.Memo = oRecordSet1.Fields.Item("Memo").Value
    '        oJV.Reference = oRecordSet1.Fields.Item("Ref1").Value
    '        oJV.Indicator = oRecordSet1.Fields.Item("Indicator").Value
    '        Dim K As Integer
    '        Dim Credit As Double = 0
    '        Dim Debit As Double = 0
    '        For i = 1 To oRecordSet1.RecordCount
    '            If i <> 1 Then
    '                oJV.Lines.Add()
    '            End If
    '            oJV.Lines.SetCurrentLine(i - 1)
    '            If oRecordSet1.Fields.Item("SIDE").Value = "Credit" Then
    '                If oRecordSet1.Fields.Item("SN").Value = "Y" Then
    '                    oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
    '                Else
    '                    oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
    '                End If
    '                Credit = Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
    '                oJV.Lines.CreditSys = Math.Round(Credit, 2)
    '                TotCredit = TotCredit + Math.Round(Credit, 2)
    '                TotCredit = Math.Round(TotCredit, 2)
    '            ElseIf oRecordSet1.Fields.Item("SIDE").Value = "Debit" Then
    '                If oRecordSet1.Fields.Item("SN").Value = "Y" Then
    '                    oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
    '                Else
    '                    oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
    '                End If
    '                If i = oRecordSet1.RecordCount Then
    '                    Debit = TotCredit - TotDebit
    '                ElseIf oRecordSet1.Fields.Item("VatGroup").Value <> "SI" Then
    '                    TotDebit = Math.Round(TotDebit, 2) + Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
    '                    Debit = oRecordSet1.Fields.Item("LOCALAMOUNT").Value
    '                Else
    '                    TotDebit = Math.Round(TotDebit, 2) + (Math.Round(TaxAmt, 2) - Math.Round(oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value, 2))
    '                    TotDebit = Math.Round(TotDebit, 2)
    '                    Debit = TaxAmt - oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value
    '                End If
    '                Debit = Math.Round(Debit, 2)
    '                oJV.Lines.DebitSys = Math.Round(Debit, 2)
    '                If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
    '                    oJV.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
    '                    oJV.Lines.SystemBaseAmount = oRecordSet1.Fields.Item("BASEAMOUNT").Value
    '                End If
    '            End If
    '            oRecordSet1.MoveNext()
    '        Next
    '        K = oJV.Add()
    '        If K <> 0 Then
    '            Dim st As String = ""
    '            ocompany.GetLastError(K, st)
    '            SBO_Application.MessageBox(st)
    '        End If
    '    Catch ex As Exception
    '        Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + " Error Message:" + ex.ToString)
    '    End Try
    'End Sub
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
