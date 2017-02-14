
Public Class F_JournalEntry
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Public Obutt As SAPbouiCOM.Button
    Public oEdit As SAPbouiCOM.EditText
    Public oCombo As SAPbouiCOM.ComboBox
    Public oItem As SAPbouiCOM.Item
    Public oNewItem As SAPbouiCOM.Item
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub
    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.FormTypeEx = 141 Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    Dim DocEntrySBO As Integer = 0
                    oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim Ojt As String = Ocompany.GetNewObjectKey()
                    oRecordSet2.DoQuery("SELECT Max(cast(""DocEntry"" as int))  FROM OPCH")
                    DocEntrySBO = oRecordSet2.Fields.Item(0).Value
                    If DocEntrySBO = (DocEtry + 1) Then
                        JournalEntry(DocEntrySBO)
                    End If
                ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = True Then
                    oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT Max(cast(""DocEntry"" as int))  FROM OPCH")
                    DocEtry = oRecordSet2.Fields.Item(0).Value
                End If
            ElseIf BusinessObjectInfo.FormTypeEx = 181 Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    Dim DocEntrySBO As Integer = 0
                    oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT Max(cast(""DocEntry"" as int))  FROM ORPC")
                    DocEntrySBO = oRecordSet2.Fields.Item(0).Value
                    If DocEntrySBO = (DocEtry + 1) Then
                        JournalEntry_CN(DocEntrySBO)
                    End If
                ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = True Then
                    oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT Max(cast(""DocEntry"" as int))  FROM ORPC")
                    DocEtry = oRecordSet2.Fields.Item(0).Value
                End If
                'ODPO -- Down Payment
            ElseIf BusinessObjectInfo.FormTypeEx = 65301 Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    Dim DocEntrySBO As Integer = 0
                    oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim Ojt As String = Ocompany.GetNewObjectKey()
                    oRecordSet2.DoQuery("SELECT Max(cast(""DocEntry"" as int))  FROM ODPO")
                    DocEntrySBO = oRecordSet2.Fields.Item(0).Value
                    If DocEntrySBO = (DocEtry + 1) Then
                        '   Exit Sub
                        JournalEntry_DP(DocEntrySBO)
                    End If
                ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = True Then
                    oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet2.DoQuery("SELECT Max(cast(""DocEntry"" as int))  FROM ODPO")
                    DocEtry = oRecordSet2.Fields.Item(0).Value
                End If
            End If
        Catch ex As Exception
            Functions.WriteLog("Class:F_JournalEntry" + " Function:SBO_Application_FormDataEvent" + " Error Message:" + ex.ToString)
        End Try

    End Sub
    Private Sub JournalEntry_CN(ByVal DocEntry As String)
        Try
            Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry_CN" + "Start Function")

            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT ""TransId"",""U_AP_EXCH_RATE"",""U_AP_TAX_AMT"",""DocRate"" FROM ORPC WHERE ""DocEntry""='" & DocEntry & "'")
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim ExRate As Double = 0
            Dim TaxAmt As Double = 0
            Dim DocRate As Double = 0
            ExRate = oRecordSet.Fields.Item("U_AP_EXCH_RATE").Value
            TaxAmt = oRecordSet.Fields.Item("U_AP_TAX_AMT").Value
            DocRate = oRecordSet.Fields.Item("DocRate").Value
            If ExRate = 0 Or TaxAmt = 0 Or DocRate = 0 Then
                Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + "ExReate:" + ExRate + "TaxAmt:" + TaxAmt + "DocRate:" + DocRate)
                Exit Sub
            End If
            'oRecordSet1.DoQuery("SELECT T0.[Project], T0.[Ref2], T0.[Ref3], T1.[Project] ProjectLine, T1.[Ref1] 'Ref1_L', T1.[Ref2] 'Ref2_L', T1.[Ref3Line], T1.[LineMemo], T1.[OcrCode2], T1.[OcrCode3], [OcrCode4], T1.[OcrCode5],T1.[ProfitCode],T0.[RefDate], T0.[DueDate], T0.[TaxDate], T0.[Memo] , 'PC' 'Indicator' ,T0.Ref1,T0.[TransId], T1.[Line_ID],CASE WHEN T1.[Account]= T1.[ShortName]  THEN 'N' ELSE 'Y' END as 'SN',T1.[ShortName] ,CASE WHEN T1.[SYSCred] <> 0 THEN 'Credit' ELSE 'Debit' END AS 'SIDE',(CASE WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCCredit]=0  THEN T1.[FCDebit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Credit]=0  THEN T1.[Debit] WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCDebit]=0  THEN T1.[FCCredit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Debit]=0  THEN T1.[Credit] END)  * " & ExRate & "- (CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END)As 'LOCALAMOUNT',CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END AS 'SYSTEMAMOUNT',CASE WHEN isnull(T1.[FCCurrency],'') <> '' THEN (((T1.[BaseSum]*" & ExRate & ")/" & DocRate & ")- T1.[SYSBaseSum]) ELSE ((T1.[BaseSum]*" & ExRate & ")- T1.[SYSBaseSum]) END AS  'BASEAMOUNT',ISNULL(T1.VatGroup,'') 'VatGroup' FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId WHERE T1.[TransId]='" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.[TransId], T1.[Line_ID] ASC")
            ' oRecordSet1.DoQuery("SELECT T0.""Project"", T0.""Ref2"", T0.""Ref3"", T1.""Project"" AS ""ProjectLine"", T1.""Ref1"" AS ""Ref1_L"", T1.""Ref2"" AS ""Ref2_L"", T1.""Ref3Line"", T1.""LineMemo"", T1.""OcrCode2"", T1.""OcrCode3"", ""OcrCode4"", T1.""OcrCode5"", T1.""ProfitCode"", T0.""RefDate"", T0.""DueDate"", T0.""TaxDate"", T0.""Memo"", 'PC' AS ""Indicator"", T0.""Ref1"", T0.""TransId"", T1.""Line_ID"", CASE WHEN T1.""Account"" = T1.""ShortName"" THEN 'N' ELSE 'Y' END AS ""SN"", T1.""ShortName"", CASE WHEN T1.""SYSCred"" <> 0 THEN 'Credit' ELSE 'Debit' END AS ""SIDE"", (CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCCredit"" = 0 THEN T1.""FCDebit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Credit"" = 0 THEN T1.""Debit"" WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCDebit"" = 0 THEN T1.""FCCredit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Debit"" = 0 THEN T1.""Credit"" END) * " & ExRate & " - (CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END) AS ""LOCALAMOUNT"", CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END AS ""SYSTEMAMOUNT"", CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' THEN (((T1.""BaseSum"" / " & ExRate & ") / " & DocRate & ") - T1.""SYSBaseSum"") ELSE ((T1.""BaseSum"" / " & ExRate & ") - T1.""SYSBaseSum"") END AS ""BASEAMOUNT"", IFNULL(T1.""VatGroup"", '') AS ""VatGroup"" FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T1.""TransId"" = '" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.""TransId"", T1.""Line_ID"" ASC")
            Dim sSQLStr As String = "SELECT T0.""Project"", T0.""Ref2"", T0.""Ref3"", T1.""Project"" AS ""ProjectLine"", T1.""Ref1"" AS ""Ref1_L"", T1.""Ref2"" AS ""Ref2_L"", T1.""Ref3Line"", T1.""LineMemo"", T1.""OcrCode2"", T1.""OcrCode3"", ""OcrCode4"", T1.""OcrCode5"", T1.""ProfitCode"", T0.""RefDate"", T0.""DueDate"", T0.""TaxDate"", T0.""Memo"", 'PU' AS ""Indicator"", T0.""Ref1"", T0.""TransId"", T1.""Line_ID"", CASE WHEN T1.""Account"" = T1.""ShortName"" THEN 'N' ELSE 'Y' END AS ""SN"", T1.""ShortName"", CASE WHEN T1.""SYSCred"" <> 0 THEN 'Credit' ELSE 'Debit' END AS ""SIDE"", (CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCCredit"" = 0 THEN T1.""FCDebit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Credit"" = 0 THEN T1.""Debit"" WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCDebit"" = 0 THEN T1.""FCCredit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Debit"" = 0 THEN T1.""Credit"" END) / " & ExRate & " - (CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END) AS ""LOCALAMOUNT"", CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END AS ""SYSTEMAMOUNT"", CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' THEN (((T1.""BaseSum"" * " & DocRate & ") / " & ExRate & ") - T1.""SYSBaseSum"") ELSE ((T1.""BaseSum"" / " & ExRate & ") - T1.""SYSBaseSum"") END AS ""BASEAMOUNT"", IFNULL(T1.""VatGroup"", '') AS ""VatGroup"" FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T1.""TransId"" = '" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.""TransId"", T1.""Line_ID"" ASC"
            oRecordSet1.DoQuery("SELECT T0.""Project"", T0.""Ref2"", T0.""Ref3"", T1.""Project"" AS ""ProjectLine"", T1.""Ref1"" AS ""Ref1_L"", T1.""Ref2"" AS ""Ref2_L"", T1.""Ref3Line"", T1.""LineMemo"", T1.""OcrCode2"", T1.""OcrCode3"", ""OcrCode4"", T1.""OcrCode5"", T1.""ProfitCode"", T0.""RefDate"", T0.""DueDate"", T0.""TaxDate"", T0.""Memo"", 'PU' AS ""Indicator"", T0.""Ref1"", T0.""TransId"", T1.""Line_ID"", CASE WHEN T1.""Account"" = T1.""ShortName"" THEN 'N' ELSE 'Y' END AS ""SN"", T1.""ShortName"", CASE WHEN T1.""SYSCred"" <> 0 THEN 'Credit' ELSE 'Debit' END AS ""SIDE"", (CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCCredit"" = 0 THEN T1.""FCDebit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Credit"" = 0 THEN T1.""Debit"" WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCDebit"" = 0 THEN T1.""FCCredit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Debit"" = 0 THEN T1.""Credit"" END) / " & ExRate & " - (CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END) AS ""LOCALAMOUNT"", CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END AS ""SYSTEMAMOUNT"", CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' THEN (((T1.""BaseSum"" * " & DocRate & ") / " & ExRate & ") - T1.""SYSBaseSum"") ELSE ((T1.""BaseSum"" / " & ExRate & ") - T1.""SYSBaseSum"") END AS ""BASEAMOUNT"", IFNULL(T1.""VatGroup"", '') AS ""VatGroup"" FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T1.""TransId"" = '" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.""TransId"", T1.""Line_ID"" ASC")
            Dim oJV As SAPbobsCOM.JournalEntries
            oJV = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            Dim i As Integer = 0
            Dim TotDebit As Double = 0
            Dim TotCredit As Double = 0
            If oRecordSet1.RecordCount = 0 Then
                Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry_CN" + "Exit Record")
                Exit Sub
            End If
            oJV.ReferenceDate = oRecordSet1.Fields.Item("RefDate").Value
            oJV.DueDate = oRecordSet1.Fields.Item("DueDate").Value
            oJV.TaxDate = oRecordSet1.Fields.Item("TaxDate").Value
            oJV.Memo = oRecordSet1.Fields.Item("Memo").Value
            oJV.Reference = oRecordSet1.Fields.Item("Ref1").Value
            oJV.Indicator = oRecordSet1.Fields.Item("Indicator").Value
            oJV.ProjectCode = oRecordSet1.Fields.Item("Project").Value
            oJV.Reference2 = oRecordSet1.Fields.Item("Ref2").Value
            oJV.Reference3 = oRecordSet1.Fields.Item("Ref3").Value

            Dim K As Integer
            Dim Credit As Double = 0
            Dim Debit As Double = 0
            For i = 1 To oRecordSet1.RecordCount
                If i <> 1 Then
                    oJV.Lines.Add()
                End If
                oJV.Lines.SetCurrentLine(i - 1)
                oJV.Lines.ProjectCode = oRecordSet1.Fields.Item("ProjectLine").Value
                oJV.Lines.Reference1 = oRecordSet1.Fields.Item("Ref1_L").Value
                oJV.Lines.Reference2 = oRecordSet1.Fields.Item("Ref2_L").Value
                '  oJV.Lines.Reference3 = oRecordSet1.Fields.Item("Ref3Line").Value

                oJV.Lines.LineMemo = oRecordSet1.Fields.Item("LineMemo").Value
                oJV.Lines.CostingCode = oRecordSet1.Fields.Item("ProfitCode").Value
                oJV.Lines.CostingCode2 = oRecordSet1.Fields.Item("OcrCode2").Value
                oJV.Lines.CostingCode3 = oRecordSet1.Fields.Item("OcrCode3").Value
                oJV.Lines.CostingCode4 = oRecordSet1.Fields.Item("OcrCode4").Value
                oJV.Lines.CostingCode5 = oRecordSet1.Fields.Item("OcrCode5").Value
                If oRecordSet1.Fields.Item("SIDE").Value = "Debit" Then 'Debit
                    If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                        oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                    Else
                        oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                    End If
                    Credit = Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                    oJV.Lines.DebitSys = Math.Round(Credit, 2)
                    TotCredit = TotCredit + Math.Round(Credit, 2)
                    TotCredit = Math.Round(TotCredit, 2)
                    If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
                        oJV.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
                        oJV.Lines.SystemBaseAmount = oRecordSet1.Fields.Item("BASEAMOUNT").Value
                    End If
                ElseIf oRecordSet1.Fields.Item("SIDE").Value = "Credit" Then
                    If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                        oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                    Else
                        oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                    End If
                    If i = oRecordSet1.RecordCount Then
                        Debit = TotCredit - TotDebit
                    ElseIf oRecordSet1.Fields.Item("VatGroup").Value <> "TX7" Then
                        TotDebit = Math.Round(TotDebit, 2) + Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                        Debit = oRecordSet1.Fields.Item("LOCALAMOUNT").Value
                    Else
                        TotDebit = Math.Round(TotDebit, 2) + (Math.Round(TaxAmt, 2) - Math.Round(oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value, 2))
                        TotDebit = Math.Round(TotDebit, 2)
                        Debit = TaxAmt - oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value
                    End If
                    Debit = Math.Round(Debit, 2)
                    oJV.Lines.CreditSys = Math.Round(Debit, 2)
                    If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
                        oJV.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
                        oJV.Lines.SystemBaseAmount = oRecordSet1.Fields.Item("BASEAMOUNT").Value
                    End If
                End If
                oRecordSet1.MoveNext()
            Next
            K = oJV.Add()
            Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry_CN" + "Start Add-Function: Return Code" & K)
            If K <> 0 Then
                Dim st As String = ""
                Ocompany.GetLastError(K, st)
                Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry_CN" + "Start Add-Function: Return Code" & st)
                If st = "Transaction Without Amount " Then
                    Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry_CN" + "Start Add-Function: Transaction Without Amount")
                    st = ""
                End If
                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("UPDATE ORPC set ""U_JEST""='N',""U_ERRMSG""='" & st & "' WHERE ""DocEntry""='" & DocEntry & "'")
            Else
                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("UPDATE ORPC set ""U_JEST""='Y' WHERE ""DocEntry""='" & DocEntry & "'")
            End If
            oJV = Nothing
            oRecordSet = Nothing
            oRecordSet1 = Nothing
            oRecordSet2 = Nothing
            GC.Collect()
        Catch ex As Exception
            Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Private Sub JournalEntry(ByVal DocEntry As String)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT ""TransId"",""U_AP_EXCH_RATE"",""U_AP_TAX_AMT"",""DocRate"" FROM OPCH WHERE ""DocEntry""='" & DocEntry & "'")
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim ExRate As Double = 0
            Dim TaxAmt As Double = 0
            Dim DocRate As Double = 0
            ExRate = oRecordSet.Fields.Item("U_AP_EXCH_RATE").Value
            TaxAmt = oRecordSet.Fields.Item("U_AP_TAX_AMT").Value
            DocRate = oRecordSet.Fields.Item("DocRate").Value
            If ExRate = 0 Or TaxAmt = 0 Or DocRate = 0 Then
                Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + "ExReate:" + ExRate + "TaxAmt:" + TaxAmt + "DocRate:" + DocRate)
                Exit Sub
            End If
            'oRecordSet1.DoQuery("SELECT T0.[Project], T0.[Ref2], T0.[Ref3], T1.[Project] ProjectLine, T1.[Ref1] 'Ref1_L', T1.[Ref2] 'Ref2_L', T1.[Ref3Line], T1.[LineMemo], T1.[OcrCode2], T1.[OcrCode3], [OcrCode4], T1.[OcrCode5],T1.[ProfitCode],T0.[RefDate], T0.[DueDate], T0.[TaxDate], T0.[Memo] , 'PU' 'Indicator' ,T0.Ref1,T0.[TransId], T1.[Line_ID],CASE WHEN T1.[Account]= T1.[ShortName]  THEN 'N' ELSE 'Y' END as 'SN',T1.[ShortName] ,CASE WHEN T1.[SYSCred] <> 0 THEN 'Credit' ELSE 'Debit' END AS 'SIDE',(CASE WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCCredit]=0  THEN T1.[FCDebit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Credit]=0  THEN T1.[Debit] WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCDebit]=0  THEN T1.[FCCredit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Debit]=0  THEN T1.[Credit] END)  * " & ExRate & "- (CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END)As 'LOCALAMOUNT',CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END AS 'SYSTEMAMOUNT',CASE WHEN isnull(T1.[FCCurrency],'') <> '' THEN (((T1.[BaseSum]*" & ExRate & ")/" & DocRate & ")- T1.[SYSBaseSum]) ELSE ((T1.[BaseSum]*" & ExRate & ")- T1.[SYSBaseSum]) END AS  'BASEAMOUNT',ISNULL(T1.VatGroup,'') 'VatGroup' FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId WHERE T1.[TransId]='" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.[TransId], T1.[Line_ID] ASC")
            oRecordSet1.DoQuery("SELECT T0.""Project"", T0.""Ref2"", T0.""Ref3"", T1.""Project"" AS ""ProjectLine"", T1.""Ref1"" AS ""Ref1_L"", T1.""Ref2"" AS ""Ref2_L"", T1.""Ref3Line"", T1.""LineMemo"", T1.""OcrCode2"", T1.""OcrCode3"", ""OcrCode4"", T1.""OcrCode5"", T1.""ProfitCode"", T0.""RefDate"", T0.""DueDate"", T0.""TaxDate"", T0.""Memo"", 'PU' AS ""Indicator"", T0.""Ref1"", T0.""TransId"", T1.""Line_ID"", CASE WHEN T1.""Account"" = T1.""ShortName"" THEN 'N' ELSE 'Y' END AS ""SN"", T1.""ShortName"", CASE WHEN T1.""SYSCred"" <> 0 THEN 'Credit' ELSE 'Debit' END AS ""SIDE"", (CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCCredit"" = 0 THEN T1.""FCDebit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Credit"" = 0 THEN T1.""Debit"" WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCDebit"" = 0 THEN T1.""FCCredit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Debit"" = 0 THEN T1.""Credit"" END) / " & ExRate & " - (CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END) AS ""LOCALAMOUNT"", CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END AS ""SYSTEMAMOUNT"", CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' THEN (((T1.""BaseSum"" * " & DocRate & ") / " & ExRate & ") - T1.""SYSBaseSum"") ELSE ((T1.""BaseSum"" / " & ExRate & ") - T1.""SYSBaseSum"") END AS ""BASEAMOUNT"", IFNULL(T1.""VatGroup"", '') AS ""VatGroup"" FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T1.""TransId"" = '" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.""TransId"", T1.""Line_ID"" ASC")
            Dim oJV As SAPbobsCOM.JournalEntries
            oJV = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            Dim i As Integer = 0
            Dim TotDebit As Double = 0
            Dim TotCredit As Double = 0
            If oRecordSet1.RecordCount = 0 Then
                Exit Sub
            End If
            oJV.ReferenceDate = oRecordSet1.Fields.Item("RefDate").Value
            oJV.DueDate = oRecordSet1.Fields.Item("DueDate").Value
            oJV.TaxDate = oRecordSet1.Fields.Item("TaxDate").Value
            oJV.Memo = oRecordSet1.Fields.Item("Memo").Value
            oJV.Reference = oRecordSet1.Fields.Item("Ref1").Value
            oJV.Indicator = oRecordSet1.Fields.Item("Indicator").Value
            oJV.ProjectCode = oRecordSet1.Fields.Item("Project").Value
            oJV.Reference2 = oRecordSet1.Fields.Item("Ref2").Value
            oJV.Reference3 = oRecordSet1.Fields.Item("Ref3").Value

            Dim DebitGSTLoc As Decimal = 0
            Dim CreditGSTLoc As Decimal = 0
            Dim DebitGSTSysc As Decimal = 0
            Dim CreditGSTSysc As Decimal = 0
            Dim DebitGSTBase As Decimal = 0
            Dim CreditGSTBase As Decimal = 0
            Dim DownPymnt As Boolean = False
            Dim K As Integer
            Dim J As Integer = 0
            Dim Credit As Double = 0
            Dim Debit As Double = 0
            Dim FC As Boolean = False
            oRecordSet1.MoveFirst()
            For i = 1 To oRecordSet1.RecordCount
                If oRecordSet1.Fields.Item("VatGroup").Value = "TX7" And oRecordSet1.Fields.Item("SIDE").Value = "Debit" Then 'Debit
                    DebitGSTLoc = Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                    DebitGSTSysc = Math.Round(oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value, 2) 'SYSTEMAMOUNT
                    DebitGSTBase = Math.Round(oRecordSet1.Fields.Item("BASEAMOUNT").Value, 2)
                ElseIf oRecordSet1.Fields.Item("VatGroup").Value = "TX7" And oRecordSet1.Fields.Item("SIDE").Value = "Credit" Then
                    DownPymnt = True
                    CreditGSTLoc = Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                    CreditGSTSysc = Math.Round(oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value, 2) 'SYSTEMAMOUNT
                    CreditGSTBase = Math.Round(oRecordSet1.Fields.Item("BASEAMOUNT").Value, 2)
                End If
                oRecordSet1.MoveNext()
            Next
            oRecordSet1.MoveFirst()
            For i = 1 To oRecordSet1.RecordCount

                If DownPymnt = False Then
                    If i <> 1 Then
                        oJV.Lines.Add()

                    End If
                    oJV.Lines.SetCurrentLine(i - 1)
                    oJV.Lines.ProjectCode = oRecordSet1.Fields.Item("ProjectLine").Value
                    oJV.Lines.Reference1 = oRecordSet1.Fields.Item("Ref1_L").Value
                    oJV.Lines.Reference2 = oRecordSet1.Fields.Item("Ref2_L").Value
                    '  oJV.Lines.Reference3 = oRecordSet1.Fields.Item("Ref3Line").Value

                    oJV.Lines.LineMemo = oRecordSet1.Fields.Item("LineMemo").Value
                    oJV.Lines.CostingCode = oRecordSet1.Fields.Item("ProfitCode").Value
                    oJV.Lines.CostingCode2 = oRecordSet1.Fields.Item("OcrCode2").Value
                    oJV.Lines.CostingCode3 = oRecordSet1.Fields.Item("OcrCode3").Value
                    oJV.Lines.CostingCode4 = oRecordSet1.Fields.Item("OcrCode4").Value
                    oJV.Lines.CostingCode5 = oRecordSet1.Fields.Item("OcrCode5").Value
                    If oRecordSet1.Fields.Item("SIDE").Value = "Credit" Then
                        'changes


                        If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                            oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                        Else
                            oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                        End If
                        Credit = Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                        oJV.Lines.CreditSys = Math.Round(Credit, 2)
                        TotCredit = TotCredit + Math.Round(Credit, 2)
                        TotCredit = Math.Round(TotCredit, 2)
                    ElseIf oRecordSet1.Fields.Item("SIDE").Value = "Debit" Then
                        If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                            oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                        Else
                            oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                        End If
                        If i = oRecordSet1.RecordCount Then
                            Debit = TotCredit - TotDebit
                        ElseIf oRecordSet1.Fields.Item("VatGroup").Value <> "TX7" Then
                            TotDebit = Math.Round(TotDebit, 2) + Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                            Debit = oRecordSet1.Fields.Item("LOCALAMOUNT").Value
                        Else
                            TotDebit = Math.Round(TotDebit, 2) + (Math.Round(TaxAmt, 2) - Math.Round(oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value, 2))
                            TotDebit = Math.Round(TotDebit, 2)
                            Debit = TaxAmt - oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value
                        End If
                        Debit = Math.Round(Debit, 2)
                        oJV.Lines.DebitSys = Math.Round(Debit, 2)
                        If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
                            oJV.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
                            oJV.Lines.SystemBaseAmount = oRecordSet1.Fields.Item("BASEAMOUNT").Value
                        End If
                    End If
                    oRecordSet1.MoveNext()
                Else

                    If i <> 1 Then
                        If FC = False Then
                            J = J + 1
                            oJV.Lines.Add()
                        End If
                    End If
                    oJV.Lines.SetCurrentLine(J)
                    oJV.Lines.ProjectCode = oRecordSet1.Fields.Item("ProjectLine").Value
                    oJV.Lines.Reference1 = oRecordSet1.Fields.Item("Ref1_L").Value
                    oJV.Lines.Reference2 = oRecordSet1.Fields.Item("Ref2_L").Value
                    '  oJV.Lines.Reference3 = oRecordSet1.Fields.Item("Ref3Line").Value

                    oJV.Lines.LineMemo = oRecordSet1.Fields.Item("LineMemo").Value
                    oJV.Lines.CostingCode = oRecordSet1.Fields.Item("ProfitCode").Value
                    oJV.Lines.CostingCode2 = oRecordSet1.Fields.Item("OcrCode2").Value
                    oJV.Lines.CostingCode3 = oRecordSet1.Fields.Item("OcrCode3").Value
                    oJV.Lines.CostingCode4 = oRecordSet1.Fields.Item("OcrCode4").Value
                    oJV.Lines.CostingCode5 = oRecordSet1.Fields.Item("OcrCode5").Value
                    FC = False
                    If oRecordSet1.Fields.Item("SIDE").Value = "Credit" Then
                        'changes
                        If oRecordSet1.Fields.Item("VatGroup").Value <> "TX7" Then
                            FC = False
                            If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                                oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                            Else
                                oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                            End If
                            Credit = Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                            oJV.Lines.CreditSys = Math.Round(Credit, 2)
                            TotCredit = TotCredit + Math.Round(Credit, 2)
                            TotCredit = Math.Round(TotCredit, 2)
                        Else
                            FC = True
                        End If
                    ElseIf oRecordSet1.Fields.Item("SIDE").Value = "Debit" Then
                        If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                            oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                        Else
                            oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                        End If
                        If i = oRecordSet1.RecordCount Then
                            Debit = TotCredit - TotDebit
                        ElseIf oRecordSet1.Fields.Item("VatGroup").Value <> "TX7" Then
                            TotDebit = Math.Round(TotDebit, 2) + Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                            Debit = oRecordSet1.Fields.Item("LOCALAMOUNT").Value
                        Else
                            TotDebit = Math.Round(TotDebit, 2) + (Math.Round(TaxAmt, 2) - (DebitGSTSysc + (CreditGSTSysc * -1))) 'Math.Round(oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value, 2))
                            TotDebit = Math.Round(TotDebit, 2)
                            Debit = TaxAmt - (DebitGSTSysc + (CreditGSTSysc * -1)) 'oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value
                        End If
                        Debit = Math.Round(Debit, 2)
                        oJV.Lines.DebitSys = Math.Round(Debit, 2)
                        If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
                            oJV.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
                            oJV.Lines.SystemBaseAmount = DebitGSTBase + (CreditGSTBase * -1) 'oRecordSet1.Fields.Item("BASEAMOUNT").Value
                        End If
                    End If
                    oRecordSet1.MoveNext()

                End If
            Next
            K = oJV.Add()
            If K <> 0 Then
                Dim st As String = ""
                Ocompany.GetLastError(K, st)
                If st = "Transaction Without Amount " Then
                    st = ""
                End If
                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("UPDATE OPCH set ""U_JEST""='N',""U_ERRMSG""='" & st.Replace("'", "''") & "' WHERE ""DocEntry""='" & DocEntry & "'")
            Else
                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("UPDATE OPCH set ""U_JEST""='Y' WHERE ""DocEntry""='" & DocEntry & "'")
            End If
            oJV = Nothing
            oRecordSet = Nothing
            oRecordSet1 = Nothing
            oRecordSet2 = Nothing
            GC.Collect()
        Catch ex As Exception
            Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Private Sub JournalEntry_DP(ByVal DocEntry As String)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT ""TransId"",""U_AP_EXCH_RATE"",""U_AP_TAX_AMT"",""DocRate"" FROM ODPO WHERE ""DocEntry""='" & DocEntry & "'")
            oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim ExRate As Double = 0
            Dim TaxAmt As Double = 0
            Dim DocRate As Double = 0

            ExRate = oRecordSet.Fields.Item("U_AP_EXCH_RATE").Value
            TaxAmt = oRecordSet.Fields.Item("U_AP_TAX_AMT").Value
            DocRate = oRecordSet.Fields.Item("DocRate").Value
            If ExRate = 0 Or TaxAmt = 0 Or DocRate = 0 Then
                Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + "ExReate:" + ExRate + "TaxAmt:" + TaxAmt + "DocRate:" + DocRate)
                Exit Sub
            End If
            ' oRecordSet1.DoQuery("SELECT T0.[Project], T0.[Ref2], T0.[Ref3], T1.[Project] ProjectLine, T1.[Ref1] 'Ref1_L', T1.[Ref2] 'Ref2_L', T1.[Ref3Line], T1.[LineMemo], T1.[OcrCode2], T1.[OcrCode3], [OcrCode4], T1.[OcrCode5],T1.[ProfitCode],T0.[RefDate], T0.[DueDate], T0.[TaxDate], T0.[Memo] , 'PU' 'Indicator' ,T0.Ref1,T0.[TransId], T1.[Line_ID],CASE WHEN T1.[Account]= T1.[ShortName]  THEN 'N' ELSE 'Y' END as 'SN',T1.[ShortName] ,CASE WHEN T1.[SYSCred] <> 0 THEN 'Credit' ELSE 'Debit' END AS 'SIDE',(CASE WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCCredit]=0  THEN T1.[FCDebit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Credit]=0  THEN T1.[Debit] WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCDebit]=0  THEN T1.[FCCredit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Debit]=0  THEN T1.[Credit] END)  * " & ExRate & "- (CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END)As 'LOCALAMOUNT',CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END AS 'SYSTEMAMOUNT',CASE WHEN isnull(T1.[FCCurrency],'') <> '' THEN (((T1.[BaseSum]*" & ExRate & ")/" & DocRate & ")- T1.[SYSBaseSum]) ELSE ((T1.[BaseSum]*" & ExRate & ")- T1.[SYSBaseSum]) END AS  'BASEAMOUNT',ISNULL(T1.VatGroup,'') 'VatGroup' FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId WHERE T1.[TransId]='" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.[TransId], T1.[Line_ID] ASC")
            'oRecordSet1.DoQuery("SELECT T0.""Project"", T0.""Ref2"", T0.""Ref3"", T1.""Project"" AS ""ProjectLine"", T1.""Ref1"" AS ""Ref1_L"", T1.""Ref2"" AS ""Ref2_L"", T1.""Ref3Line"", T1.""LineMemo"", T1.""OcrCode2"", T1.""OcrCode3"", ""OcrCode4"", T1.""OcrCode5"", T1.""ProfitCode"", T0.""RefDate"", T0.""DueDate"", T0.""TaxDate"", T0.""Memo"", 'PU' AS ""Indicator"", T0.""Ref1"", T0.""TransId"", T1.""Line_ID"", CASE WHEN T1.""Account"" = T1.""ShortName"" THEN 'N' ELSE 'Y' END AS ""SN"", T1.""ShortName"", CASE WHEN T1.""SYSCred"" <> 0 THEN 'Credit' ELSE 'Debit' END AS ""SIDE"", (CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCCredit"" = 0 THEN T1.""FCDebit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Credit"" = 0 THEN T1.""Debit"" WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCDebit"" = 0 THEN T1.""FCCredit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Debit"" = 0 THEN T1.""Credit"" END) * " & ExRate & " - (CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END) AS ""LOCALAMOUNT"", CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END AS ""SYSTEMAMOUNT"", CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' THEN (((T1.""BaseSum"" / " & ExRate & ") / " & DocRate & ") - T1.""SYSBaseSum"") ELSE ((T1.""BaseSum"" / " & ExRate & ") - T1.""SYSBaseSum"") END AS ""BASEAMOUNT"", IFNULL(T1.""VatGroup"", '') AS ""VatGroup"" FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T1.""TransId"" = '" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.""TransId"", T1.""Line_ID"" ASC")
            Dim SQStr As String = "SELECT T0.""Project"", T0.""Ref2"", T0.""Ref3"", T1.""Project"" AS ""ProjectLine"", T1.""Ref1"" AS ""Ref1_L"", T1.""Ref2"" AS ""Ref2_L"", T1.""Ref3Line"", T1.""LineMemo"", T1.""OcrCode2"", T1.""OcrCode3"", ""OcrCode4"", T1.""OcrCode5"", T1.""ProfitCode"", T0.""RefDate"", T0.""DueDate"", T0.""TaxDate"", T0.""Memo"", 'PU' AS ""Indicator"", T0.""Ref1"", T0.""TransId"", T1.""Line_ID"", CASE WHEN T1.""Account"" = T1.""ShortName"" THEN 'N' ELSE 'Y' END AS ""SN"", T1.""ShortName"", CASE WHEN T1.""SYSCred"" <> 0 THEN 'Credit' ELSE 'Debit' END AS ""SIDE"", (CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCCredit"" = 0 THEN T1.""FCDebit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Credit"" = 0 THEN T1.""Debit"" WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCDebit"" = 0 THEN T1.""FCCredit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Debit"" = 0 THEN T1.""Credit"" END) / " & ExRate & " - (CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END) AS ""LOCALAMOUNT"", CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END AS ""SYSTEMAMOUNT"", CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' THEN (((T1.""BaseSum"" * " & DocRate & ") / " & ExRate & ") - T1.""SYSBaseSum"") ELSE ((T1.""BaseSum"" / " & ExRate & ") - T1.""SYSBaseSum"") END AS ""BASEAMOUNT"", IFNULL(T1.""VatGroup"", '') AS ""VatGroup"" FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T1.""TransId"" = '" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.""TransId"", T1.""Line_ID"" ASC"
            oRecordSet1.DoQuery("SELECT T0.""Project"", T0.""Ref2"", T0.""Ref3"", T1.""Project"" AS ""ProjectLine"", T1.""Ref1"" AS ""Ref1_L"", T1.""Ref2"" AS ""Ref2_L"", T1.""Ref3Line"", T1.""LineMemo"", T1.""OcrCode2"", T1.""OcrCode3"", ""OcrCode4"", T1.""OcrCode5"", T1.""ProfitCode"", T0.""RefDate"", T0.""DueDate"", T0.""TaxDate"", T0.""Memo"", 'PU' AS ""Indicator"", T0.""Ref1"", T0.""TransId"", T1.""Line_ID"", CASE WHEN T1.""Account"" = T1.""ShortName"" THEN 'N' ELSE 'Y' END AS ""SN"", T1.""ShortName"", CASE WHEN T1.""SYSCred"" <> 0 THEN 'Credit' ELSE 'Debit' END AS ""SIDE"", (CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCCredit"" = 0 THEN T1.""FCDebit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Credit"" = 0 THEN T1.""Debit"" WHEN IFNULL(T1.""FCCurrency"", '') <> '' AND T1.""FCDebit"" = 0 THEN T1.""FCCredit"" WHEN IFNULL(T1.""FCCurrency"", '') = '' AND T1.""Debit"" = 0 THEN T1.""Credit"" END) / " & ExRate & " - (CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END) AS ""LOCALAMOUNT"", CASE WHEN T1.""SYSCred"" = 0 THEN T1.""SYSDeb"" ELSE T1.""SYSCred"" END AS ""SYSTEMAMOUNT"", CASE WHEN IFNULL(T1.""FCCurrency"", '') <> '' THEN (((T1.""BaseSum"" * " & DocRate & ") / " & ExRate & ") - T1.""SYSBaseSum"") ELSE ((T1.""BaseSum"" / " & ExRate & ") - T1.""SYSBaseSum"") END AS ""BASEAMOUNT"", IFNULL(T1.""VatGroup"", '') AS ""VatGroup"" FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T1.""TransId"" = '" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.""TransId"", T1.""Line_ID"" ASC")
            Dim oJV As SAPbobsCOM.JournalEntries
            oJV = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            Dim i As Integer = 0
            Dim TotDebit As Double = 0
            Dim TotCredit As Double = 0
            If oRecordSet1.RecordCount = 0 Then
                Exit Sub
            End If
            oJV.ReferenceDate = oRecordSet1.Fields.Item("RefDate").Value
            oJV.DueDate = oRecordSet1.Fields.Item("DueDate").Value
            oJV.TaxDate = oRecordSet1.Fields.Item("TaxDate").Value
            oJV.Memo = oRecordSet1.Fields.Item("Memo").Value
            oJV.Reference = oRecordSet1.Fields.Item("Ref1").Value
            oJV.Indicator = oRecordSet1.Fields.Item("Indicator").Value
            oJV.ProjectCode = oRecordSet1.Fields.Item("Project").Value
            oJV.Reference2 = oRecordSet1.Fields.Item("Ref2").Value
            oJV.Reference3 = oRecordSet1.Fields.Item("Ref3").Value

            Dim K As Integer
            Dim Credit As Double = 0
            Dim Debit As Double = 0
            For i = 1 To oRecordSet1.RecordCount
                If i <> 1 Then
                    oJV.Lines.Add()
                End If
                oJV.Lines.SetCurrentLine(i - 1)
                oJV.Lines.ProjectCode = oRecordSet1.Fields.Item("ProjectLine").Value
                oJV.Lines.Reference1 = oRecordSet1.Fields.Item("Ref1_L").Value
                oJV.Lines.Reference2 = oRecordSet1.Fields.Item("Ref2_L").Value
                '  oJV.Lines.Reference3 = oRecordSet1.Fields.Item("Ref3Line").Value

                oJV.Lines.LineMemo = oRecordSet1.Fields.Item("LineMemo").Value
                oJV.Lines.CostingCode = oRecordSet1.Fields.Item("ProfitCode").Value
                oJV.Lines.CostingCode2 = oRecordSet1.Fields.Item("OcrCode2").Value
                oJV.Lines.CostingCode3 = oRecordSet1.Fields.Item("OcrCode3").Value
                oJV.Lines.CostingCode4 = oRecordSet1.Fields.Item("OcrCode4").Value
                oJV.Lines.CostingCode5 = oRecordSet1.Fields.Item("OcrCode5").Value
                If oRecordSet1.Fields.Item("SIDE").Value = "Credit" Then
                    If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                        oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                    Else
                        oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                    End If
                    Credit = Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                    oJV.Lines.CreditSys = Math.Round(Credit, 2)
                    TotCredit = TotCredit + Math.Round(Credit, 2)
                    TotCredit = Math.Round(TotCredit, 2)
                ElseIf oRecordSet1.Fields.Item("SIDE").Value = "Debit" Then
                    If oRecordSet1.Fields.Item("SN").Value = "Y" Then
                        oJV.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
                    Else
                        oJV.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
                    End If
                    If i = oRecordSet1.RecordCount Then
                        Debit = TotCredit - TotDebit
                    ElseIf oRecordSet1.Fields.Item("VatGroup").Value <> "TX7" Then
                        TotDebit = Math.Round(TotDebit, 2) + Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
                        Debit = oRecordSet1.Fields.Item("LOCALAMOUNT").Value
                    Else
                        TotDebit = Math.Round(TotDebit, 2) + (Math.Round(TaxAmt, 2) - Math.Round(oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value, 2))
                        TotDebit = Math.Round(TotDebit, 2)
                        Debit = TaxAmt - oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value
                    End If
                    Debit = Math.Round(Debit, 2)
                    oJV.Lines.DebitSys = Math.Round(Debit, 2)
                    If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
                        oJV.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
                        oJV.Lines.SystemBaseAmount = oRecordSet1.Fields.Item("BASEAMOUNT").Value
                    End If
                End If
                oRecordSet1.MoveNext()
            Next
            K = oJV.Add()
            If K <> 0 Then
                Dim st As String = ""
                Ocompany.GetLastError(K, st)
                If st = "Transaction Without Amount " Then
                    st = ""
                End If
                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("UPDATE ODPO set ""U_JEST""='N',""U_ERRMSG""='" & st & "' WHERE ""DocEntry""='" & DocEntry & "'")
            Else
                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("UPDATE ODPO set ""U_JEST""='Y' WHERE ""DocEntry""='" & DocEntry & "'")
            End If
            oJV = Nothing
            oRecordSet = Nothing
            oRecordSet1 = Nothing
            oRecordSet2 = Nothing
            GC.Collect()
        Catch ex As Exception
            Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry_DP" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    'Private Sub JournalEntry(ByVal DocEntry As String)
    '    Try
    '        oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRecordSet.DoQuery("SELECT [TransId],U_AP_EXCH_RATE,U_AP_TAX_AMT FROM OPCH WHERE DocEntry='" & DocEntry & "'")
    '        oRecordSet1 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim ExRate As Double = 0
    '        Dim TaxAmt As Double = 0
    '        ExRate = oRecordSet.Fields.Item("U_AP_EXCH_RATE").Value
    '        TaxAmt = oRecordSet.Fields.Item("U_AP_TAX_AMT").Value
    '        oRecordSet1.DoQuery("SELECT T0.Ref1,T0.[TransId], T1.[Line_ID],CASE WHEN T1.[Account]= T1.[ShortName]  THEN 'N' ELSE 'Y' END as 'SN',T1.[ShortName] ,CASE WHEN T1.[SYSCred] <> 0 THEN 'Credit' ELSE 'Debit' END AS 'SIDE',(CASE WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCCredit]=0  THEN T1.[FCDebit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Credit]=0  THEN T1.[Debit] WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCDebit]=0  THEN T1.[FCCredit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Debit]=0  THEN T1.[Credit] END)  * " & ExRate & "- (CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END)As 'LOCALAMOUNT',CASE WHEN T1.[SYSCred] = 0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END AS 'SYSTEMAMOUNT',((T1.[BaseSum]*" & ExRate & ")- T1.[SYSBaseSum])'BASEAMOUNT',ISNULL(T1.VatGroup,'') 'VatGroup' FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId WHERE T1.[TransId]='" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.[TransId], T1.[Line_ID] ASC")

    '        '  Dim oJV As SAPbobsCOM.JournalEntries = DirectCast(Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries), SAPbobsCOM.JournalEntries)
    '        Dim oJV As SAPbobsCOM.JournalEntries
    '        oJV = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

    '        'Dim oJV As SAPbobsCOM.IJournalEntries
    '        'oJV = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    '        Dim i As Integer = 0
    '        Dim TotDebit As Double = 0
    '        Dim TotCredit As Double = 0
    '        If oRecordSet1.RecordCount = 0 Then
    '            Exit Sub
    '        End If



    '        oJV.JournalEntries.Reference = "10" 'oRecordSet1.Fields.Item("Ref1").Value.ToString
    '        oJV.JournalEntries.Indicator = "PU"
    '        For i = 1 To oRecordSet1.RecordCount
    '            If oRecordSet1.Fields.Item("SIDE").Value = "Credit" Then
    '                If oRecordSet1.Fields.Item("SN").Value = "Y" Then
    '                    oJV.JournalEntries.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
    '                Else
    '                    oJV.JournalEntries.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
    '                End If
    '                oJV.JournalEntries.Lines.CreditSys = oRecordSet1.Fields.Item("LOCALAMOUNT").Value
    '                'If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
    '                '    oJV.JournalEntries.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
    '                '    oJV.JournalEntries.Lines.SystemBaseAmount = oRecordSet1.Fields.Item("BASEAMOUNT").Value
    '                'End If
    '                TotCredit = TotCredit + oRecordSet1.Fields.Item("LOCALAMOUNT").Value
    '                TotCredit = Math.Round(TotCredit, 2)
    '                oJV.JournalEntries.Lines.Add()
    '            ElseIf oRecordSet1.Fields.Item("SIDE").Value = "Debit" Then
    '                If oRecordSet1.Fields.Item("SN").Value = "Y" Then
    '                    oJV.JournalEntries.Lines.ShortName = oRecordSet1.Fields.Item("ShortName").Value
    '                Else
    '                    oJV.JournalEntries.Lines.AccountCode = oRecordSet1.Fields.Item("ShortName").Value
    '                End If

    '                If i = oRecordSet1.RecordCount Then
    '                    oJV.JournalEntries.Lines.DebitSys = TotCredit - TotDebit
    '                ElseIf oRecordSet1.Fields.Item("VatGroup").Value <> "TX7" Then
    '                    TotDebit = Math.Round(TotDebit, 2) + Math.Round(oRecordSet1.Fields.Item("LOCALAMOUNT").Value, 2)
    '                    oJV.JournalEntries.Lines.DebitSys = oRecordSet1.Fields.Item("LOCALAMOUNT").Value
    '                Else
    '                    TotDebit = Math.Round(TotDebit, 2) + (Math.Round(TaxAmt, 2) - oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value)
    '                    TotDebit = Math.Round(TotDebit, 2)
    '                    oJV.JournalEntries.Lines.DebitSys = TaxAmt - oRecordSet1.Fields.Item("SYSTEMAMOUNT").Value
    '                End If
    '                If oRecordSet1.Fields.Item("VatGroup").Value <> "" Then
    '                    oJV.JournalEntries.Lines.TaxGroup = oRecordSet1.Fields.Item("VatGroup").Value
    '                    oJV.JournalEntries.Lines.SystemBaseAmount = oRecordSet1.Fields.Item("BASEAMOUNT").Value
    '                End If
    '                oJV.JournalEntries.Lines.Add()
    '            End If
    '            oRecordSet1.MoveNext()
    '        Next

    '        Dim K As Integer
    '        k = oJV.Add()

    '        If K <> 0 Then
    '            Dim st As String = ""
    '            Ocompany.GetLastError(K, st)
    '            SBO_Application.MessageBox(st)
    '        End If


    '    Catch ex As Exception
    '        Functions.WriteLog("Class:F_JournalEntry" + " Function:JournalEntry" + " Error Message:" + ex.ToString)
    '    End Try
    'End Sub
    ' oRecordSet1.DoQuery("SELECT T0.[TransId], T1.[Line_ID],CASE WHEN T1.[Account]= T1.[ShortName]  THEN 'N' ELSE'Y' END as 'ShortName',T1.[ShortName] ,CASE WHEN T1.[SYSCred] <> 0 THEN 'Credit' ELSE 'Debit' END AS 'SIDE',CASE WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCCredit]=0 THEN T1.[FCDebit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Credit]=0 THEN T1.[Debit] WHEN isnull(T1.[FCCurrency],'') <> '' AND T1.[FCDebit] =0 THEN T1.[FCCredit] WHEN isnull(T1.[FCCurrency],'') = '' AND T1.[Debit] =0 THEN T1.[Credit] END As 'LOCAL AMOUNT',CASE WHEN [SYSCred]=0 THEN T1.[SYSDeb] ELSE T1.[SYSCred] END AS 'SYSTEM AMOUNT',T1.[BaseSum], T1.[SYSBaseSum],ISNULL(T1.VatGroup,'') 'VatGroup' FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId WHERE T1.[TransId] ='" & oRecordSet.Fields.Item(0).Value.ToString & "' ORDER BY T0.[TransId], T1.[Line_ID] ASC")

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormType = 141 Then
                If pVal.ItemUID = "GSR141" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    'oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, 141)
                    Try
                        Dim DocEntrySBO As Integer = 0
                        oEdit = oForm.Items.Item("8").Specific
                        Dim docNum As String = ""
                        docNum = oEdit.String
                        If docNum <> "" Then
                            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet2.DoQuery("SELECT cast(""DocEntry"" as int) as ""DocEntry"", ""U_JEST""  FROM OPCH where ""DocNum"" = '" & docNum & "' ")
                            If oRecordSet2.Fields.Item("U_JEST").Value.ToString <> "Y" Then
                                DocEntrySBO = oRecordSet2.Fields.Item("DocEntry").Value
                                JournalEntry(DocEntrySBO)
                            End If
                        End If
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
                oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then
                    Try
                        oItem = oForm.Items.Item("2")
                        oNewItem = oForm.Items.Add("GSR141", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + (oItem.Width * 2)
                        oNewItem.Width = oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = oItem.Height
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "GST Retry"
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If

            If pVal.FormType = 181 Then
                If pVal.ItemUID = "GSR181" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False And pVal.InnerEvent = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    'oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, 141)
                    Try
                        Dim DocEntrySBO As Integer = 0
                        oEdit = oForm.Items.Item("8").Specific
                        Dim docNum As String = ""
                        docNum = oEdit.String
                        If docNum <> "" Then
                            oRecordSet2 = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet2.DoQuery("SELECT cast(""DocEntry"" as int) as ""DocEntry"", ""U_JEST""  FROM ORPC where ""DocNum"" = '" & docNum & "' ")
                            If oRecordSet2.Fields.Item("U_JEST").Value.ToString <> "Y" Then
                                DocEntrySBO = oRecordSet2.Fields.Item("DocEntry").Value
                                JournalEntry_CN(DocEntrySBO)
                            End If
                        End If
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If

                oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.Before_Action = True)) Then
                    Try
                        oItem = oForm.Items.Item("2")
                        oNewItem = oForm.Items.Add("GSR181", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oNewItem.Left = oItem.Left + (oItem.Width * 2)
                        oNewItem.Width = oItem.Width
                        oNewItem.Top = oItem.Top
                        oNewItem.Height = oItem.Height
                        Obutt = oNewItem.Specific
                        Obutt.Caption = "GST Retry"
                    Catch ex As Exception
                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If
        Catch ex As Exception
        End Try


    End Sub
End Class
