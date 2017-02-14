Module modMain
#Region "Company Default"
    Public Structure CompanyDefault

        Public sServer As String
        Public sSourceDBName As String
        Public sTargetDBName As String
        Public sDBUser As String
        Public sDBPwd As String
        Public sSourceSAPUser As String
        Public sTargetSAPUser As String
        Public sSourceSAPPwd As String
        Public sTargetSAPPwd As String
        Public sDriver As String

        Public sDebug As String
        Public sFilepath As String

    End Structure
#End Region

#Region "Global Variable"
    ' Return Value Variable Control
    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16 = DEBUG_ON
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault
    Public p_dProcessing As DateTime
    Public p_oDtSuccess As DataTable
    Public p_oDtError As DataTable
    Public p_SyncDateTime As String
    Public p_SrcCompany As SAPbobsCOM.Company
    Public p_oCompany As SAPbobsCOM.Company
    Public oTrgIMasterCompany As New SAPbobsCOM.Company
#End Region

    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty

        Dim strConnection As Array
        Dim sApp As String = String.Empty

        ''For ItemMaster
        Dim sItemCode As String = String.Empty
        Dim sICode As String = String.Empty

        Dim oDT_ItemLog As DataTable = Nothing
        Dim oDTS_ItemLog As DataTable = Nothing
        Dim oDT_ItemSync As DataTable = Nothing
        Dim oDT_ItemCode As DataTable = Nothing

        Dim oRset_Item As SAPbobsCOM.Recordset = Nothing
        Dim oRset_ItemUpdate As SAPbobsCOM.Recordset = Nothing

        Dim sIUpdateSyncSQL As String = String.Empty
        Dim sItemSyncSQl As String = String.Empty
        Dim sUpdateItemSQL As String = String.Empty


        ''For BP master
        Dim sBPCode As String = String.Empty
        Dim sVCode As String = String.Empty

        Dim oDT_BPLog As DataTable = Nothing
        Dim oDTS_BPLog As DataTable = Nothing
        Dim oDT_BPSync As DataTable = Nothing
        Dim oDT_BPCode As DataTable = Nothing

        Dim oRset_BP As SAPbobsCOM.Recordset = Nothing
        Dim oRset_BPUpdate As SAPbobsCOM.Recordset = Nothing

        Dim sBPUpdateSyncSQL As String = String.Empty
        Dim sBPSyncSQl As String = String.Empty
        Dim sUpdateBPSQL As String = String.Empty

        ''For BP Price List
        Dim sBPPLCode As String = String.Empty
        Dim sPLItemCode As String = String.Empty
        Dim sPCode As String = String.Empty

        Dim oDT_BPPLLog As DataTable = Nothing
        Dim oDTS_BPPLLog As DataTable = Nothing
        Dim oDT_BPPLSync As DataTable = Nothing
        Dim oDT_BPPLCode As DataTable = Nothing

        Dim oRset_BPPL As SAPbobsCOM.Recordset = Nothing
        Dim oRset_BPPLUpdate As SAPbobsCOM.Recordset = Nothing

        Dim sBPPLUpdateSyncSQL As String = String.Empty
        Dim sBPPLSyncSQl As String = String.Empty
        Dim sUpdateBPPLSQL As String = String.Empty

        Try
            sFuncName = "Main'"
            Console.WriteLine("Starting Main Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            p_SrcCompany = New SAPbobsCOM.Company
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            sApp = p_oCompDef.sSourceDBName & ";" & p_oCompDef.sTargetDBName & ";" & p_oCompDef.sServer & ";" & p_oCompDef.sDBUser & ";" & p_oCompDef.sDBPwd & ";" & p_oCompDef.sSourceSAPUser & ";" & p_oCompDef.sSourceSAPPwd & ";" & p_oCompDef.sTargetSAPUser & ";" & p_oCompDef.sTargetSAPPwd
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("App " & sApp, sFuncName)
            strConnection = sApp.Split(";")

            ' '''******************************************Connection  of Source & Target Company Started ******************************************
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ConnectToCompany()", sFuncName)
            If (ConnectToCompany(strConnection, p_SrcCompany, sErrDesc) <> 0) Then
                Throw New Exception(sErrDesc)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Source Connected SUCCESSFULLY ", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ConnectToTargetCompany()", sFuncName)
            'Dim oTrgIMasterCompany As New SAPbobsCOM.Company
            If (ConnectToTargetCompany(strConnection, oTrgIMasterCompany, sErrDesc) <> 0) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Call WriteToLogFile(sErrDesc, sFuncName)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target Connected SUCCESSFULLY ", sFuncName)
            ' '''******************************************Connection  of Source & Target Company Ended ******************************************

            ' '''******************************************Item Master Sync Starts******************************************
            Try

                oDT_ItemLog = New DataTable()
                oDT_ItemLog.Columns.Add("Code", GetType(String))
                oDT_ItemLog.Columns.Add("Status", GetType(String))
                oDT_ItemLog.Columns.Add("Msg", GetType(String))

                'oDTS_ItemLog = New DataTable()
                'oDTS_ItemLog.Columns.Add("Code", GetType(String))
                'oDTS_ItemLog.Columns.Add("Status", GetType(String))
                'oDTS_ItemLog.Columns.Add("Msg", GetType(String))

                sItemSyncSQl = "SELECT T0.""Code"", T0.""Name"", T0.""ItemCode"" FROM ""AE_ITEMMASTER_SYNC""  T0 WHERE T0.""Status""  <> 'SYNCHRONIZED'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ITEM  SYNC " & sItemSyncSQl, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HANAtoDatatable() ", sFuncName)
                oDT_ItemCode = HANAtoDatatable(sItemSyncSQl, sErrDesc)
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)
                oRset_Item = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_ItemUpdate = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


                If oDT_ItemCode.Rows.Count > 0 Then
                    'For Each odr As DataRow In oDT_ItemCode.Rows
                    '    sICode += "'" & odr("Code") & "',"
                    'Next
                    'sICode = Left(sICode, sICode.Length - 1)
                    For Each oItemCode As DataRow In oDT_ItemCode.Rows
                        sItemCode = oItemCode("ItemCode").ToString.Trim()
                        If UpdateItemMaster(sItemCode, p_SrcCompany, oTrgIMasterCompany, sErrDesc) <> RTN_SUCCESS Then
                            Console.WriteLine("ItemCode - " & sItemCode & " - is been failed while Synchronizing", sFuncName)
                            WriteToLogFile_Sync(sItemCode.PadRight(20, " "c) & "  " & Now.ToLongDateString.PadRight(30, " "c) & "FAIL   " & "  " & sErrDesc)
                            '' sErrDesc = sErrDesc.Remove("'")
                            oDT_ItemLog.Rows.Add(oItemCode("Code").ToString.Trim(), "FAILED", sErrDesc)
                            sErrDesc = ""
                        Else
                            oDT_ItemLog.Rows.Add(oItemCode("Code").ToString.Trim(), "SYNCHRONIZED", sErrDesc)
                            Console.WriteLine("ItemCode - " & sItemCode & " - is Synchronized Successfully", sFuncName)
                            WriteToLogFile_Sync(sItemCode.PadRight(20, " "c) & "  " & Now.ToLongDateString.PadRight(30, " "c) & "SUCCESS" & "  " & sErrDesc)
                            sUpdateItemSQL = "update ""OITM"" set ""U_SyncStatus"" = 'Synchronized' where ""ItemCode"" in ('" & sItemCode & "')"
                            oRset_ItemUpdate.DoQuery(sUpdateItemSQL)
                            sUpdateItemSQL = ""
                        End If
                    Next

                    'sIUpdateSyncSQL = "update ""@AE_ITEMMASTER_SYNC"" set ""U_Status"" = 'Synchronized', ""U_SyncDate"" = CURRENT_DATE, ""U_SyncTime"" = '" & Now.ToShortTimeString & "' , ""U_ErrMsg"" = '' where ""Code"" in (" & sICode & ")"
                    'oRset_Item.DoQuery(sIUpdateSyncSQL)
                    'sIUpdateSyncSQL = ""
                    If oDT_ItemLog.Rows.Count > 0 Then
                        For Each olog As DataRow In oDT_ItemLog.Rows
                            'sIUpdateSyncSQL = "update ""AE_ITEMMASTER_SYNC""  set ""Status"" = 'FAILED', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '" & olog("Msg").ToString.Replace("'", "''") & "' where ""Code"" = '" & olog("Code") & "'"
                            sIUpdateSyncSQL = "update ""AE_ITEMMASTER_SYNC""  set ""Status"" = '" & olog("Status") & "', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '" & olog("Msg").ToString.Replace("'", "''") & "' where ""Code"" = '" & olog("Code") & "'"
                            oRset_Item.DoQuery(sIUpdateSyncSQL)
                            sIUpdateSyncSQL = ""
                        Next

                    End If

                    'If oDTS_ItemLog.Rows.Count > 0 Then
                    '    For Each olog As DataRow In oDTS_ItemLog.Rows
                    '        'sIUpdateSyncSQL = "update ""AE_ITEMMASTER_SYNC"" set ""Status"" = 'Synchronized', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '' where ""Code"" in (" & sICode & ")"
                    '        sIUpdateSyncSQL = "update ""AE_ITEMMASTER_SYNC"" set ""Status"" = 'Synchronized', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '' where ""Code"" = '" & olog("Code") & "'"
                    '        oRset_Item.DoQuery(sIUpdateSyncSQL)
                    '        sIUpdateSyncSQL = ""

                    '    Next
                    '    Console.WriteLine("All or Some ItemCode for ItemMaster Sync Completed With SUCCESS ", sFuncName)
                    'End If
                    Console.WriteLine("Item Master Synchronization is completed", sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No ItemMaster Information for SYNC ", sFuncName)
                    Console.WriteLine("No ItemMaster Information for Synchronization", sFuncName)
                End If
            Catch ex As Exception
                Console.WriteLine("Item Master Synchronization Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Master Sync Completed with ERROR", sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
            End Try
            ''''******************************************Item Master Sync Ends******************************************


            ''''******************************************BP Master Sync Starts******************************************
            Try

                oDT_BPLog = New DataTable()
                oDT_BPLog.Columns.Add("Code", GetType(String))
                oDT_BPLog.Columns.Add("Status", GetType(String))
                oDT_BPLog.Columns.Add("Msg", GetType(String))

                'oDTS_BPLog = New DataTable()
                'oDTS_BPLog.Columns.Add("Code", GetType(String))
                'oDTS_BPLog.Columns.Add("Status", GetType(String))
                'oDTS_BPLog.Columns.Add("Msg", GetType(String))

                sBPSyncSQl = "SELECT T0.""Code"", T0.""Name"", T0.""BPCode"" FROM ""AE_BP_SYNC""  T0 WHERE T0.""Status""  <> 'SYNCHRONIZED'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP  SYNC " & sBPSyncSQl, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HANAtoDatatable() ", sFuncName)
                oDT_BPCode = HANAtoDatatable(sBPSyncSQl, sErrDesc)
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)

                oRset_BP = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_BPUpdate = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


                If oDT_BPCode.Rows.Count > 0 Then
                    'For Each odr As DataRow In oDT_BPCode.Rows
                    '    sVCode += "'" & odr("Code") & "',"
                    'Next
                    'sVCode = Left(sVCode, sVCode.Length - 1)

                    For Each oBPCode As DataRow In oDT_BPCode.Rows
                        sBPCode = oBPCode("BPCode").ToString.Trim()
                        If UpdateBPMaster(sBPCode, p_SrcCompany, oTrgIMasterCompany, sErrDesc) <> RTN_SUCCESS Then
                            WriteToLogFile_Sync(sBPCode.PadRight(20, " "c) & "  " & Now.ToLongDateString.PadRight(30, " "c) & "FAIL   " & "  " & sErrDesc)
                            Console.WriteLine("BPCode - " & sBPCode & " - is been failed while Synchronizing", sFuncName)
                            oDT_BPLog.Rows.Add(oBPCode("Code").ToString.Trim(), "FAILED", sErrDesc)
                            sErrDesc = ""
                        Else
                            oDT_BPLog.Rows.Add(oBPCode("Code").ToString.Trim(), "SYNCHRONIZED", "")
                            WriteToLogFile_Sync(sItemCode.PadRight(20, " "c) & "  " & Now.ToLongDateString.PadRight(30, " "c) & "SUCCESS" & "  " & sErrDesc)
                            Console.WriteLine("BPCode - " & sBPCode & " - is Synchronized Successfully", sFuncName)
                            sUpdateBPSQL = "update ""OCRD"" set ""U_SyncStatus"" = 'Synchronized' where ""CardCode"" in ('" & sBPCode & "')"
                            oRset_BPUpdate.DoQuery(sUpdateBPSQL)
                            sUpdateBPSQL = ""
                        End If
                    Next

                  
                    If oDT_BPLog.Rows.Count > 0 Then
                        For Each olog As DataRow In oDT_BPLog.Rows
                            'sBPUpdateSyncSQL = "update ""AE_BP_SYNC""  set ""Status"" = 'FAILED', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '" & olog("Msg").ToString.Replace("'", "''") & "' where ""Code"" = '" & olog("Code") & "'"
                            sBPUpdateSyncSQL = "update ""AE_BP_SYNC""  set ""Status"" = '" & olog("Status") & "', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '" & olog("Msg").ToString.Replace("'", "''") & "' where ""Code"" = '" & olog("Code") & "'"
                            oRset_BP.DoQuery(sBPUpdateSyncSQL)
                            sBPUpdateSyncSQL = ""
                        Next                       
                    End If
                    Console.WriteLine("BP Master Synchronization is  Completed", sFuncName)
                    'If oDTS_BPLog.Rows.Count > 0 Then
                    '    'sBPUpdateSyncSQL = "update ""AE_BP_SYNC"" set ""Status"" = 'Synchronized', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '' where ""Code"" in (" & sVCode & ")"
                    '    For Each olog As DataRow In oDTS_BPLog.Rows
                    '        sBPUpdateSyncSQL = "update ""AE_BP_SYNC"" set ""Status"" = 'Synchronized', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '' where ""Code""  = '" & olog("Code") & "'"
                    '        oRset_BP.DoQuery(sBPUpdateSyncSQL)
                    '        sBPUpdateSyncSQL = ""
                    '    Next

                    '    Console.WriteLine("All or Some BPCode for BP Master Sync Completed With SUCCESS ", sFuncName)
                    'End If
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No BPMaster Information for SYNC ", sFuncName)
                    Console.WriteLine("No BPMaster Information for Synchronization  ", sFuncName)
                End If

            Catch ex As Exception
                Console.WriteLine("BP Master Synchronization Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP Master Sync Completed with ERROR", sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
            End Try
            ''''******************************************BP Master Sync Ends********************************************

            ''''******************************************BP Price List Sync Starts********************************************
            Try


                oDT_BPPLLog = New DataTable()
                oDT_BPPLLog.Columns.Add("Code", GetType(String))
                oDT_BPPLLog.Columns.Add("Status", GetType(String))
                oDT_BPPLLog.Columns.Add("Msg", GetType(String))

                'oDTS_BPPLLog = New DataTable()
                'oDTS_BPPLLog.Columns.Add("Code", GetType(String))
                'oDTS_BPPLLog.Columns.Add("Status", GetType(String))
                'oDTS_BPPLLog.Columns.Add("Msg", GetType(String))

                sBPPLSyncSQl = "SELECT T0.""Code"", T0.""Name"", T0.""BPCode"", T0.""ItemCode"" FROM ""AE_BP_PRICE_SYNC""  T0 WHERE T0.""Status""  <> 'SYNCHRONIZED'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP Price List  SYNC " & sBPPLSyncSQl, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HANAtoDatatable() ", sFuncName)
                oDT_BPPLCode = HANAtoDatatable(sBPPLSyncSQl, sErrDesc)
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)

                oRset_BPPL = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_BPPLUpdate = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


                If oDT_BPPLCode.Rows.Count > 0 Then
                    For Each odr As DataRow In oDT_BPPLCode.Rows
                        sPCode += "'" & odr("Code") & "',"
                    Next
                    sPCode = Left(sPCode, sPCode.Length - 1)

                    For Each oBPPLCode As DataRow In oDT_BPPLCode.Rows
                        sBPPLCode = oBPPLCode("BPCode").ToString.Trim()
                        sPLItemCode = oBPPLCode("ItemCode").ToString.Trim()

                        If UpdateBPPriceList(sBPPLCode, sPLItemCode, p_SrcCompany, oTrgIMasterCompany, sErrDesc) <> RTN_SUCCESS Then
                            WriteToLogFile_Sync(sBPPLCode.PadRight(20, " "c) & "  " & Now.ToLongDateString.PadRight(30, " "c) & "FAIL   " & "  " & sErrDesc)
                            Console.WriteLine("BP Price List - " & sBPPLCode & " - " & sPLItemCode & " - is been failed while Synchronizing", sFuncName)
                            'sErrDesc = sErrDesc.Remove("'")
                            '' sErrDesc = sErrDesc.Replace("'", """")
                            oDT_BPPLLog.Rows.Add(oBPPLCode("Code").ToString.Trim(), "FAILED", sErrDesc)
                            sErrDesc = ""
                        Else
                            oDT_BPPLLog.Rows.Add(oBPPLCode("Code").ToString.Trim(), "SYNCHRONIZED", "")
                            Console.WriteLine("BP Price List - " & sBPPLCode & " - " & sPLItemCode & " - is Synchronized Successfully", sFuncName)
                            WriteToLogFile_Sync(sBPPLCode.PadRight(20, " "c) & "  " & Now.ToLongDateString.PadRight(30, " "c) & "SUCCESS" & "  " & sErrDesc)
                            sUpdateBPPLSQL = "update ""OSPP"" set ""U_SyncStatus"" = 'Synchronized' where ""CardCode"" in ('" & sBPPLCode & "') and ""ItemCode"" in ('" & sPLItemCode & "') "
                            oRset_BPUpdate.DoQuery(sUpdateBPPLSQL)
                            sUpdateBPPLSQL = ""
                        End If
                    Next

                   
                    If oDT_BPPLLog.Rows.Count > 0 Then
                        For Each olog As DataRow In oDT_BPPLLog.Rows
                            'sBPPLUpdateSyncSQL = "update ""AE_BP_PRICE_SYNC""  set ""Status"" = 'FAILED', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '" & olog("Msg").ToString.Replace("'", "''") & "' where ""Code"" = '" & olog("Code") & "'"
                            sBPPLUpdateSyncSQL = "update ""AE_BP_PRICE_SYNC""  set ""Status"" = '" & olog("Status") & "', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '" & olog("Msg").ToString.Replace("'", "''") & "' where ""Code"" = '" & olog("Code") & "'"
                            oRset_BPPL.DoQuery(sBPPLUpdateSyncSQL)
                            sBPPLUpdateSyncSQL = ""
                        Next                                               
                    End If
                    Console.WriteLine("BP Price List Synchronization is  Completed", sFuncName)
                    'If oDTS_BPPLLog.Rows.Count > 0 Then
                    '    For Each olog As DataRow In oDTS_BPPLLog.Rows
                    '        'sBPPLUpdateSyncSQL = "update ""AE_BP_PRICE_SYNC"" set ""Status"" = 'Synchronized', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '' where ""Code"" in (" & sPCode & ")"
                    '        sBPPLUpdateSyncSQL = "update ""AE_BP_PRICE_SYNC"" set ""Status"" = 'Synchronized', ""SyncDate"" = CURRENT_DATE, ""SyncTime"" = '" & Now.ToShortTimeString & "' , ""ErrMsg"" = '' where ""Code""  = '" & olog("Code") & "'"
                    '        oRset_BPPL.DoQuery(sBPPLUpdateSyncSQL)
                    '        sBPPLUpdateSyncSQL = ""
                    '    Next

                    '    Console.WriteLine("All or Some of BP Price List Sync Completed With SUCCESS ", sFuncName)
                    'End If
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No BP Price List Information for Synchronization ", sFuncName)
                    Console.WriteLine("No BP Price List Information for Synchronization  ", sFuncName)
                End If

            Catch ex As Exception
                Console.WriteLine("BP Price List Synchronization Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP Price List Sync Completed with ERROR", sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
            End Try

            ''''******************************************BP Price List Sync Ends**********************************************


        Catch ex As Exception
            Console.WriteLine("Synchronization Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sync Completed with ERROR", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
      

        End Try

    End Sub



End Module
