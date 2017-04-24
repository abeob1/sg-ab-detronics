Public Class F_JournalEntry
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.FormTypeEx = 141 Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    Dim docEntry As String = ""
                    Ocompany.GetNewObjectCode(docEntry)
                    JournalEntry(docEntry)
                End If
            End If
        Catch ex As Exception
            Functions.WriteLog("Class:F_JournalEntry" + " Function:SBO_Application_FormDataEvent" + " Error Message:" + ex.ToString)
        End Try

    End Sub
    Private Sub JournalEntry(ByVal DocEntry As String)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT [TransId] FROM OPCH WHERE DocEntry='" & DocEntry & "'")
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet1.DoQuery("SELECT T0.[TransId], T1.[Line_ID],CASE WHEN T1.[Account]= T1.[ShortName]  THEN 'N' ELSE'Y' END as 'ShortName',T1.[ShortName] ,CASE WHEN T1.[SYSCred] <> 0 THEN 'Credit' ELSE 'Debit' END AS 'SIDE',CASE WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCCredit]=0 THEN T1.[FCDebit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Credit]=0 THEN T1.[Debit] WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCDebit] =0 THEN T1.[FCCredit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Debit] =0 THEN T1.[Credit] END As 'LOCAL AMOUNT',CASE WHEN [SYSCred]=0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END AS 'SYSTEM AMOUNT',T1.[BaseSum], T1.[SYSBaseSum],ISNULL(T1.VatGroup,'') 'VatGroup' FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId WHERE T1.[TransId] ='" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.[TransId], T1.[Line_ID] ASC")

        Catch ex As Exception
            Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + " Error Message:" + ex.ToString)
        End Try
    End Sub
End Class
