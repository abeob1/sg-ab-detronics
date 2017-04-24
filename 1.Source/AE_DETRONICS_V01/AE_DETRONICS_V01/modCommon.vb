Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Data.Odbc
Imports System.Data.Common
Imports Sap.Data.Hana

Module modCommon

#Region "GetSystemIntializeInfo"
    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing information about the initialing variables
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   Shibin
        '   Date        :   SEP 2016
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSqlstr As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            Console.WriteLine("Starting System Intial  Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sFuncName = "GetSystemIntializeInfo()"
            ' Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)


            oCompDef.sServer = String.Empty
            oCompDef.sDBUser = String.Empty
            oCompDef.sDBPwd = String.Empty
            oCompDef.sSourceDBName = String.Empty
            oCompDef.sTargetDBName = String.Empty
            oCompDef.sDebug = String.Empty

            'PublicVariable.SourceConnection = System.Configuration.ConfigurationManager.ConnectionStrings("SourceConnection").ConnectionString

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SERVERNODE").ToString) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("SERVERNODE").ToString
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("UID").ToString) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("UID").ToString
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PWD").ToString) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("PWD").ToString
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SOURCECS").ToString) Then
                oCompDef.sSourceDBName = ConfigurationManager.AppSettings("SOURCECS").ToString
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TARGETCS").ToString) Then
                oCompDef.sTargetDBName = ConfigurationManager.AppSettings("TARGETCS").ToString
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SOURCESAPUser").ToString) Then
                oCompDef.sSourceSAPUser = ConfigurationManager.AppSettings("SOURCESAPUser").ToString
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SOURCESAPPWD").ToString) Then
                oCompDef.sSourceSAPPwd = ConfigurationManager.AppSettings("SOURCESAPPWD").ToString
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TARGETSAPUser").ToString) Then
                oCompDef.sTargetSAPUser = ConfigurationManager.AppSettings("TARGETSAPUser").ToString
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TARGETSAPPWD").ToString) Then
                oCompDef.sTargetSAPPwd = ConfigurationManager.AppSettings("TARGETSAPPWD").ToString
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DRIVER").ToString) Then
                oCompDef.sDriver = ConfigurationManager.AppSettings("DRIVER").ToString
            End If

            ' folder

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sFilepath = ConfigurationManager.AppSettings("LogPath")
            End If

            Console.WriteLine("System Intial is Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("System Intial is Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function
#End Region

#Region "Connect To Company"
    Public Function ConnectToCompany(ByVal Connection As Array, ByVal company As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectToCompany()
        '   Purpose     :   This function will connect to Source Company
        '               
        '   Parameters  :   ByVal Connection As Array
        '                       Connection =  set the Connection String
        '                   ByVal company As SAPbobsCOM.Company
        '                       company = set the Company
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   Shibin
        '   Date        :   SEP 2016
        ' **********************************************************************************
        Dim sErrMsg As String = ""
        Dim sErrCode As Integer
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "ConnectToCompany()"
            If company.Connected Then
                company.Disconnect()
            End If


            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            company.CompanyDB = Connection(0).ToString()
            company.Server = Connection(2).ToString()
            company.DbUserName = Connection(3).ToString()
            company.DbPassword = Connection(4).ToString()
            company.UserName = Connection(5).ToString
            company.Password = Connection(6).ToString

            sErrDesc = String.Empty

            If company.Connect <> 0 Then
                company.GetLastError(sErrCode, sErrMsg)
                sErrDesc = sErrMsg
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Source Company not connected " & Connection(0).ToString() & " - " & sErrMsg, "ConnectToCompany()")
                Console.WriteLine("Source Company not connected ", sFuncName)
            Else
                Console.WriteLine("Source Company connected  successfully ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Source Company connected  successfully " & Connection(0).ToString(), "ConnectToCompany()")
            End If
            Return sErrCode
        Catch ex As Exception
            'WriteLog("SystemInitial: " + ex.ToString)
            Return ex.ToString
        End Try

    End Function

    Public Function ConnectToTargetCompany(ByVal Connection As Array, ByVal company As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   ConnectToTargetCompany()
        '   Purpose     :   This function will connect to Source Company
        '               
        '   Parameters  :   ByVal Connection As Array
        '                       Connection =  set the Connection String
        '                   ByVal company As SAPbobsCOM.Company
        '                       company = set the Company
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   Shibin
        '   Date        :   SEP 2016
        ' **********************************************************************************
        Dim sErrMsg As String = ""
        Dim sErrCode As Integer
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "ConnectToTargetCompany()"
            If company.Connected Then
                company.Disconnect()
            End If


            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            company.CompanyDB = Connection(1).ToString()
            company.Server = Connection(2).ToString()
            company.DbUserName = Connection(3).ToString()
            company.DbPassword = Connection(4).ToString()
            company.UserName = Connection(7).ToString
            company.Password = Connection(8).ToString

            sErrDesc = String.Empty

            If company.Connect <> 0 Then
                company.GetLastError(sErrCode, sErrMsg)
                sErrDesc = sErrMsg
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target Company not connected " & Connection(1).ToString() & " - " & sErrMsg, "ConnectToCompany()")
                Console.WriteLine("Target Company not connected", sFuncName)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target Company connected  successfully " & Connection(1).ToString(), "ConnectToCompany()")
                Console.WriteLine("Target Company connected  successfully", sFuncName)
            End If
            Return sErrCode
        Catch ex As Exception
            'WriteLog("SystemInitial: " + ex.ToString)
            Return ex.ToString
        End Try

    End Function
#End Region

#Region "Execute HANA Query- Datatable"
    Public Function HANAtoDatatable(ByVal sQuery As String, ByRef sErrDesc As String) As DataTable
        '  **********************************************************************************
        '    Function    :   HANAtoDatatable()
        '    Purpose     :   This function will fetch the information based on the query and fill the Datatable
        '                
        '    Parameters  :  
        '                    ByRef sErrDesc AS String 
        '                        sErrDesc = Error Description to be returned to calling function
        '                
        '    Return      :   0 - FAILURE
        '                    1 - SUCCESS
        '    Author      :   Shibin
        '    Date        :   Sep 2016
        '  *********************************************************************************


        Dim sFuncName As String = "HANAtoDatatable"
        Dim oDataset As New DataSet()
        'Dim sConnString As String = System.Configuration.ConfigurationManager.ConnectionStrings("SourceHanaConnection").ConnectionString
        'Dim oHanaOdbcConnection As New OdbcConnection(sConnString)
        'Dim oHanaConnection As HanaConnection = New HanaConnection(sConnString)

        Dim sConstr As String = "DRIVER=" & p_oCompDef.sDriver & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSourceDBName
        Dim oHanaOdbcConnection As New OdbcConnection(sConstr)
        Dim oHanaOdbcCommand As New OdbcCommand()
        Dim oHanaConnection As HanaConnection = New HanaConnection(sConstr)
        Try
            sFuncName = "HANAtoDatatable()"

            'Console.WriteLine("Starting Hana Query ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("HANA Connection", sFuncName)

            If oHanaConnection.State = ConnectionState.Closed Then
                oHanaConnection.Open()
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully HANA Connection done and Query passed..........", sFuncName)
            Dim cmd As HanaCommand = New HanaCommand(sQuery, oHanaConnection)
            'Dim reader As HanaDataReader = cmd.ExecuteReader()
            Dim oHanaDA As New HanaDataAdapter(cmd)
            oHanaDA.Fill(oDataset)
            oHanaDA.Dispose()
            'reader.Close()
            cmd.Dispose()

            Return oDataset.Tables(0)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Query retreived result..........", sFuncName)

        Catch Ex As Exception
            sErrDesc = Ex.Message
            'Console.WriteLine("Completed with ERROR ", sFuncName)
            WriteToLogFile(Ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in HANA Connection/Query passed..........", sFuncName)
            Throw New Exception(Ex.Message)
        Finally
            oHanaConnection.Close()
            oHanaConnection.Dispose()
        End Try
    End Function
#End Region

#Region "Item Master"
    Public Function UpdateItemMaster(ByVal ItemCode As String, ByRef oCompany As SAPbobsCOM.Company, ByRef targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim oItemMaster As SAPbobsCOM.Items = Nothing
        Dim oTItemMaster As SAPbobsCOM.Items = Nothing
        Dim orsGroup As SAPbobsCOM.Recordset = Nothing
        Dim GroupName As String = ""
        Dim sFuncName As String = String.Empty
        sFuncName = "UpdateItemMaster()"
        Dim sSQL As String = String.Empty

        oItemMaster = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
        'ItemCode = "ItemTest1"
        If oItemMaster.GetByKey(Left(ItemCode, 50)) Then

            Try
                Dim sErrMsg As String = ""
                Dim sErrCode As Integer = 0
                
                If targetCompany.Connected Then
                    Dim bfound As Boolean = False
                    oTItemMaster = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Conected with target Company " & targetCompany.CompanyDB, sFuncName)

                    If oTItemMaster.GetByKey(Left(ItemCode, 50)) Then
                        oTItemMaster.ItemName = oItemMaster.ItemName
                        oTItemMaster.ItemType = oItemMaster.ItemType
                        oTItemMaster.ForeignName = oItemMaster.ForeignName

                        orsGroup = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""ItmsGrpNam"" from ""OITB"" where ""ItmsGrpCod"" = {0}", oItemMaster.ItemsGroupCode))
                        GroupName = orsGroup.Fields.Item(0).Value
                        orsGroup = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""ItmsGrpCod"" from ""OITB"" where ""ItmsGrpNam"" = '{0}'", GroupName))
                        If orsGroup.RecordCount = 1 Then
                            oTItemMaster.ItemsGroupCode = orsGroup.Fields.Item(0).Value
                        End If

                        'oTItemMaster.ItemsGroupCode = oItemMaster.ItemsGroupCode

                        oTItemMaster.InventoryItem = oItemMaster.InventoryItem
                        oTItemMaster.SalesItem = oItemMaster.SalesItem
                        oTItemMaster.PurchaseItem = oItemMaster.PurchaseItem
                        oTItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                        'oTItemMaster.PurchaseVATGroup = oItemMaster.PurchaseVATGroup
                        oTItemMaster.PurchaseVATGroup = "NAI"
                        oTItemMaster.GLMethod = oItemMaster.GLMethod
                        oTItemMaster.WTLiable = oItemMaster.WTLiable
                        oTItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit

                        'oTItemMaster.PreferredVendors = oItemMaster.PreferredVendors
                        oTItemMaster.SupplierCatalogNo = oItemMaster.SupplierCatalogNo

                        oTItemMaster.Manufacturer = oItemMaster.Manufacturer
                        oTItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit
                        oTItemMaster.PurchasePackagingUnit = oItemMaster.PurchasePackagingUnit
                        oTItemMaster.PurchaseQtyPerPackUnit = oItemMaster.PurchaseQtyPerPackUnit
                        oTItemMaster.PurchaseItemsPerUnit = oItemMaster.PurchaseItemsPerUnit


                        'oTItemMaster.SalesVATGroup = oItemMaster.SalesVATGroup
                        oTItemMaster.SalesVATGroup = "NAO"
                        oTItemMaster.SalesUnit = oItemMaster.SalesUnit
                        oTItemMaster.SalesPackagingUnit = oItemMaster.SalesPackagingUnit
                        oTItemMaster.SalesQtyPerPackUnit = oItemMaster.SalesQtyPerPackUnit
                        oTItemMaster.SalesItemsPerUnit = oItemMaster.SalesItemsPerUnit

                        oTItemMaster.InventoryUoMEntry = oItemMaster.InventoryUoMEntry
                        'oTItemMaster.OrderIntervals = oItemMaster.OrderIntervals
                        'oTItemMaster.QuantityOrderedFromVendors = oItemMaster.QuantityOrderedFromVendors
                        'oTItemMaster.QuantityOrderedByCustomers = oItemMaster.QuantityOrderedByCustomers
                        oTItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                        oTItemMaster.MaxInventory = oItemMaster.MaxInventory
                        oTItemMaster.MinInventory = oItemMaster.MinInventory
                        oTItemMaster.MinOrderQuantity = oItemMaster.MinOrderQuantity


                        'If oTItemMaster.WhsInfo.Count > 0 Then
                        '    Dim delete As Boolean = False
                        '    For i As Integer = 0 To oTItemMaster.WhsInfo.Count - 1
                        '        oTItemMaster.WhsInfo.SetCurrentLine(oTItemMaster.WhsInfo.Count - 1)
                        '        oTItemMaster.WhsInfo.Delete()
                        '        If oTItemMaster.WhsInfo.Count = 0 Then
                        '            Exit For
                        '        End If
                        '    Next
                        'End If

                        'For iLine As Integer = 0 To oItemMaster.WhsInfo.Count - 1
                        '    oItemMaster.WhsInfo.SetCurrentLine(iLine)
                        '    oTItemMaster.WhsInfo.WarehouseCode = oItemMaster.WhsInfo.WarehouseCode
                        '    oTItemMaster.WhsInfo.ExpensesAccount = oItemMaster.WhsInfo.ExpensesAccount
                        '    oTItemMaster.WhsInfo.ForeignExpensAcc = oItemMaster.WhsInfo.ForeignExpensAcc
                        '    oTItemMaster.WhsInfo.PurchaseCreditAcc = oItemMaster.WhsInfo.PurchaseCreditAcc
                        '    oTItemMaster.WhsInfo.ForeignPurchaseCreditAcc = oItemMaster.WhsInfo.ForeignPurchaseCreditAcc
                        '    oTItemMaster.WhsInfo.Add()
                        'Next

                        oTItemMaster.Employee = oItemMaster.Employee
                        oTItemMaster.Properties(1) = oItemMaster.Properties(1)
                        oTItemMaster.Properties(2) = oItemMaster.Properties(2)
                        oTItemMaster.Properties(3) = oItemMaster.Properties(3)
                        oTItemMaster.Properties(4) = oItemMaster.Properties(4)
                        oTItemMaster.Properties(5) = oItemMaster.Properties(5)
                        oTItemMaster.Properties(6) = oItemMaster.Properties(6)
                        oTItemMaster.Properties(7) = oItemMaster.Properties(7)
                        oTItemMaster.Properties(8) = oItemMaster.Properties(8)
                        oTItemMaster.Properties(9) = oItemMaster.Properties(9)
                        oTItemMaster.Properties(10) = oItemMaster.Properties(10)
                        oTItemMaster.Properties(11) = oItemMaster.Properties(11)
                        oTItemMaster.Properties(12) = oItemMaster.Properties(12)

                        oTItemMaster.User_Text = oItemMaster.User_Text
                        oTItemMaster.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                        oTItemMaster.Valid = SAPbobsCOM.BoYesNoEnum.tYES

                        oTItemMaster.FrozenFrom = oItemMaster.FrozenFrom
                        oTItemMaster.FrozenTo = oItemMaster.FrozenTo
                        oTItemMaster.ValidFrom = oItemMaster.ValidFrom
                        oTItemMaster.ValidTo = oItemMaster.ValidTo

                        oTItemMaster.UserFields.Fields.Item("U_COO").Value = oItemMaster.UserFields.Fields.Item("U_COO").Value
                        oTItemMaster.UserFields.Fields.Item("U_GeneralCode").Value = oItemMaster.UserFields.Fields.Item("U_GeneralCode").Value
                        oTItemMaster.UserFields.Fields.Item("U_QtyPerCarton_PC").Value = oItemMaster.UserFields.Fields.Item("U_QtyPerCarton_PC").Value
                        oTItemMaster.UserFields.Fields.Item("U_ProductGrp").Value = oItemMaster.UserFields.Fields.Item("U_ProductGrp").Value
                        oTItemMaster.UserFields.Fields.Item("U_WtPerCarton").Value = oItemMaster.UserFields.Fields.Item("U_WtPerCarton").Value


                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating ItemCode " & ItemCode, sFuncName)
                        If oTItemMaster.Update() <> 0 Then
                            targetCompany.GetLastError(sErrCode, sErrMsg)
                            Throw New Exception("Could not update ItemCode to Target Company" + " - " + sErrMsg)
                            'Console.WriteLine("Could not update ItemCode to Target Company - " & ItemCode, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & "Could not update ItemCode to Target Company" + " - " + sErrMsg, sFuncName)
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                            'Console.WriteLine("Updated ItemCode to Target Company - " & ItemCode, sFuncName)
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function CreateItemMaster()", sFuncName)
                        If CreateItemMaster(ItemCode, oTItemMaster, oItemMaster, targetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    End If
                End If
                UpdateItemMaster = RTN_SUCCESS
                sErrDesc = String.Empty

            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                UpdateItemMaster = RTN_ERROR
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItemMaster)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTItemMaster)
                oItemMaster = Nothing
                oTItemMaster = Nothing
            End Try
        Else
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error: ItemCode not found!!!!", sFuncName)
            UpdateItemMaster = RTN_ERROR
            sErrDesc = "ItemCode not found"
        End If

    End Function

    Public Function CreateItemMaster(ByVal ItemCode As String, ByVal oTItemMaster As SAPbobsCOM.Items, ByVal oItemMaster As SAPbobsCOM.Items, ByVal targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName = "CreateItemMaster()"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
        Dim orsGroup As SAPbobsCOM.Recordset = Nothing

        Dim GroupName As String = ""
        Dim sTrgtShipTypeCode As String = String.Empty
        Dim sSrcShipTypeName As String = String.Empty
        If oItemMaster.GetByKey(ItemCode) Then
            Try
                If oTItemMaster.GetByKey(oItemMaster.ItemCode) Then
                    Throw New Exception(String.Format("ItemCode : {0} aldready existed in Target Company : {1}", oItemMaster.ItemCode, targetCompany.CompanyName))
                End If
                oTItemMaster.ItemCode = oItemMaster.ItemCode
                oTItemMaster.ItemName = oItemMaster.ItemName
                oTItemMaster.ForeignName = oItemMaster.ForeignName
                oTItemMaster.ItemType = oItemMaster.ItemType

                orsGroup = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orsGroup.DoQuery(String.Format("Select ""ItmsGrpNam"" from ""OITB"" where ""ItmsGrpCod"" = {0}", oItemMaster.ItemsGroupCode))
                GroupName = orsGroup.Fields.Item(0).Value
                orsGroup = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orsGroup.DoQuery(String.Format("Select ""ItmsGrpCod"" from ""OITB"" where ""ItmsGrpNam"" = '{0}'", GroupName))
                If orsGroup.RecordCount = 1 Then
                    oTItemMaster.ItemsGroupCode = orsGroup.Fields.Item(0).Value
                End If

                'oTItemMaster.ItemsGroupCode = oItemMaster.ItemsGroupCode
                oTItemMaster.InventoryItem = oItemMaster.InventoryItem
                oTItemMaster.SalesItem = oItemMaster.SalesItem
                oTItemMaster.PurchaseItem = oItemMaster.PurchaseItem
                oTItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                'oTItemMaster.PurchaseVATGroup = oItemMaster.PurchaseVATGroup
                oTItemMaster.PurchaseVATGroup = "NAI"
                oTItemMaster.GLMethod = oItemMaster.GLMethod
                oTItemMaster.WTLiable = oItemMaster.WTLiable
                oTItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit

                'oTItemMaster.PreferredVendors = oItemMaster.PreferredVendors
                oTItemMaster.SupplierCatalogNo = oItemMaster.SupplierCatalogNo
                oTItemMaster.Manufacturer = oItemMaster.Manufacturer
                oTItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit
                oTItemMaster.PurchasePackagingUnit = oItemMaster.PurchasePackagingUnit
                oTItemMaster.PurchaseQtyPerPackUnit = oItemMaster.PurchaseQtyPerPackUnit
                oTItemMaster.PurchaseItemsPerUnit = oItemMaster.PurchaseItemsPerUnit

                'oTItemMaster.SalesVATGroup = oItemMaster.SalesVATGroup
                oTItemMaster.SalesVATGroup = "NAO"
                oTItemMaster.SalesUnit = oItemMaster.SalesUnit
                oTItemMaster.SalesPackagingUnit = oItemMaster.SalesPackagingUnit
                oTItemMaster.SalesQtyPerPackUnit = oItemMaster.SalesQtyPerPackUnit
                oTItemMaster.SalesItemsPerUnit = oItemMaster.SalesItemsPerUnit


                oTItemMaster.InventoryUoMEntry = oItemMaster.InventoryUoMEntry
                'oTItemMaster.OrderIntervals = oItemMaster.OrderIntervals
                'oTItemMaster.QuantityOrderedFromVendors = oItemMaster.QuantityOrderedFromVendors
                'oTItemMaster.QuantityOrderedByCustomers = oItemMaster.QuantityOrderedByCustomers
                oTItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                oTItemMaster.MaxInventory = oItemMaster.MaxInventory
                oTItemMaster.MinInventory = oItemMaster.MinInventory
                oTItemMaster.MinOrderQuantity = oItemMaster.MinOrderQuantity


                For iLine As Integer = 0 To oItemMaster.WhsInfo.Count - 1
                    oItemMaster.WhsInfo.SetCurrentLine(iLine)
                    oTItemMaster.WhsInfo.WarehouseCode = oItemMaster.WhsInfo.WarehouseCode
                    oTItemMaster.WhsInfo.ExpensesAccount = oItemMaster.WhsInfo.ExpensesAccount
                    oTItemMaster.WhsInfo.ForeignExpensAcc = oItemMaster.WhsInfo.ForeignExpensAcc
                    oTItemMaster.WhsInfo.PurchaseCreditAcc = oItemMaster.WhsInfo.PurchaseCreditAcc
                    oTItemMaster.WhsInfo.ForeignPurchaseCreditAcc = oItemMaster.WhsInfo.ForeignPurchaseCreditAcc
                    oTItemMaster.WhsInfo.Add()
                Next
                oTItemMaster.Employee = oItemMaster.Employee
                oTItemMaster.Properties(1) = oItemMaster.Properties(1)
                oTItemMaster.Properties(2) = oItemMaster.Properties(2)
                oTItemMaster.Properties(3) = oItemMaster.Properties(3)
                oTItemMaster.Properties(4) = oItemMaster.Properties(4)
                oTItemMaster.Properties(5) = oItemMaster.Properties(5)
                oTItemMaster.Properties(6) = oItemMaster.Properties(6)
                oTItemMaster.Properties(7) = oItemMaster.Properties(7)
                oTItemMaster.Properties(8) = oItemMaster.Properties(8)
                oTItemMaster.Properties(9) = oItemMaster.Properties(9)
                oTItemMaster.Properties(10) = oItemMaster.Properties(10)
                oTItemMaster.Properties(11) = oItemMaster.Properties(11)
                oTItemMaster.Properties(12) = oItemMaster.Properties(12)

                oTItemMaster.User_Text = oItemMaster.User_Text
                oTItemMaster.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                oTItemMaster.Valid = SAPbobsCOM.BoYesNoEnum.tYES

                oTItemMaster.FrozenFrom = oItemMaster.FrozenFrom
                oTItemMaster.FrozenTo = oItemMaster.FrozenTo
                oTItemMaster.ValidFrom = oItemMaster.ValidFrom
                oTItemMaster.ValidTo = oItemMaster.ValidTo

                oTItemMaster.UserFields.Fields.Item("U_COO").Value = oItemMaster.UserFields.Fields.Item("U_COO").Value
                oTItemMaster.UserFields.Fields.Item("U_GeneralCode").Value = oItemMaster.UserFields.Fields.Item("U_GeneralCode").Value
                oTItemMaster.UserFields.Fields.Item("U_QtyPerCarton_PC").Value = oItemMaster.UserFields.Fields.Item("U_QtyPerCarton_PC").Value
                oTItemMaster.UserFields.Fields.Item("U_ProductGrp").Value = oItemMaster.UserFields.Fields.Item("U_ProductGrp").Value
                oTItemMaster.UserFields.Fields.Item("U_WtPerCarton").Value = oItemMaster.UserFields.Fields.Item("U_WtPerCarton").Value


                If oTItemMaster.Add() <> 0 Then
                    Dim errCode As Integer
                    Dim errMess As String = ""
                    targetCompany.GetLastError(errCode, errMess)
                    Throw New Exception("Could not create ItemCode to Target Company- " & ItemCode + " - " + errMess)
                    'Console.WriteLine("Could not create ItemCode to Target Company - " & ItemCode, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & "Could not create ItemCode to Target Company- " & ItemCode + " - " + errMess, sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                    'Console.WriteLine("Created ItemCode to Target Company- " & ItemCode, sFuncName)
                End If

                CreateItemMaster = RTN_SUCCESS
                sErrDesc = String.Empty

            Catch ex As Exception
                sErrDesc = ex.Message
                '' oMail.Dispose()
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                CreateItemMaster = RTN_ERROR
            End Try
        End If
    End Function
#End Region

#Region "BP Master"
    Public Function UpdateBPMaster(ByVal CardCode As String, ByRef oCompany As SAPbobsCOM.Company, ByRef targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
        Dim orsB As SAPbobsCOM.Recordset = Nothing
        Dim orsTarget As SAPbobsCOM.Recordset = Nothing
        Dim orsGroup As SAPbobsCOM.Recordset = Nothing
        Dim oTargetBP As SAPbobsCOM.BusinessPartners = Nothing
        Dim GroupName As String = ""
        Dim sFuncName As String = String.Empty
        sFuncName = "UpdateBPMaster()"
        Dim sSQL As String = String.Empty
        Dim oDVContact As DataView = Nothing
        Dim sSrcShipTypeName As String = String.Empty
        Dim sSrcPymntTrmsCod As String = String.Empty

        oBP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        'If oBP.GetByKey(Left(CardCode, 30)) Then
        If oBP.GetByKey(CardCode) Then
            Try
                Dim sErrMsg As String = ""
                Dim sErrCode As Integer = 0

                If targetCompany.Connected Then
                    Dim bfound As Boolean = False
                    orsB = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orsTarget = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTargetBP = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    sSQL = "select ROW_NUMBER() OVER(ORDER BY T1.""CntctCode"" ) -1 ""No"", T1.""CntctCode"" , T1.""Name"" , T1.""Position""  from" & _
                          """OCPR"" T1  where T1.""CardCode"" ='" & CardCode & "' order by T1.""CntctCode"" "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Contact Info " & sSQL, sFuncName)
                    orsB.DoQuery(sSQL)
                    oDVContact = New DataView(ConvertRecordset(orsB, sErrDesc))
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Conected with Company " & targetCompany.CompanyDB, sFuncName)
                    If oTargetBP.GetByKey(CardCode) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update BP " & CardCode, sFuncName)

                        oTargetBP.CardName = oBP.CardName
                        oTargetBP.CardType = oBP.CardType
                        oTargetBP.CardForeignName = oBP.CardForeignName
                        oTargetBP.CompanyPrivate = oBP.CompanyPrivate
                        oTargetBP.DiscountPercent = oBP.DiscountPercent
                        oTargetBP.Address = oBP.Address
                        oTargetBP.EmailAddress = oBP.EmailAddress
                        oTargetBP.Phone1 = oBP.Phone1
                        oTargetBP.Phone2 = oBP.Phone2
                        oTargetBP.Cellular = oBP.Cellular
                        oTargetBP.Fax = oBP.Fax
                        oTargetBP.Password = oBP.Password
                        oTargetBP.BusinessType = oBP.BusinessType
                        oTargetBP.AdditionalID = oBP.AdditionalID
                        oTargetBP.VatIDNum = oBP.VatIDNum
                        oTargetBP.FederalTaxID = oBP.FederalTaxID
                        oTargetBP.Notes = oBP.Notes
                        oTargetBP.FreeText = oBP.FreeText
                        oTargetBP.AliasName = oBP.AliasName
                        oTargetBP.GlobalLocationNumber = oBP.GlobalLocationNumber
                        oTargetBP.Valid = oBP.Valid
                        oTargetBP.Frozen = oBP.Frozen

                        oTargetBP.Website = oBP.Website
                        oTargetBP.UnifiedFederalTaxID = oBP.UnifiedFederalTaxID

                        orsGroup = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""GroupName"" from ""OCRG"" where ""GroupCode"" = {0}", oBP.GroupCode))
                        GroupName = orsGroup.Fields.Item(0).Value

                        orsGroup = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '{0}'", GroupName))
                        If orsGroup.RecordCount = 1 Then
                            oTargetBP.GroupCode = orsGroup.Fields.Item(0).Value
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping Code", sFuncName)
                        orsGroup = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        sSQL = "SELECT ""TrnspName"" FROM ""OSHP"" WHERE ""TrnspCode"" = (SELECT ""ShipType"" FROM ""OCRD"" WHERE ""CardCode"" = '" & oBP.CardCode & "')"
                        orsGroup.DoQuery(sSQL)
                        If orsGroup.RecordCount = 1 Then
                            sSrcShipTypeName = orsGroup.Fields.Item(0).Value
                        End If
                        If sSrcShipTypeName <> "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                            orsGroup = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orsGroup.DoQuery(String.Format("SELECT ""TrnspCode"" FROM ""OSHP"" WHERE ""TrnspName"" = '{0}'", sSrcShipTypeName))
                            If orsGroup.RecordCount = 1 Then
                                oTargetBP.ShippingType = orsGroup.Fields.Item(0).Value
                            End If
                        End If

                        If oTargetBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.ValidFrom = oBP.ValidFrom
                            oTargetBP.ValidTo = oBP.ValidTo
                            oTargetBP.ValidRemarks = oBP.ValidRemarks
                        End If
                        If oTargetBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.FrozenFrom = oBP.FrozenFrom
                            oTargetBP.FrozenTo = oBP.FrozenTo
                            oTargetBP.FrozenRemarks = oBP.FrozenRemarks
                        End If
                        If oTargetBP.Addresses.Count > 0 Then

                            Dim delete As Boolean = False
                            For i As Integer = 0 To oTargetBP.Addresses.Count - 1
                                oTargetBP.Addresses.SetCurrentLine(oTargetBP.Addresses.Count - 1)
                                oTargetBP.Addresses.Delete()
                                If oTargetBP.Addresses.Count = 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        'Handle add update/add new Address
                        If oBP.Addresses.Count > 0 And oBP.Addresses.AddressName <> "" Then
                            For i As Integer = 0 To oBP.Addresses.Count - 1
                                oBP.Addresses.SetCurrentLine(i)
                                oTargetBP.Addresses.AddressName = oBP.Addresses.AddressName
                                oTargetBP.Addresses.AddressName2 = oBP.Addresses.AddressName2
                                oTargetBP.Addresses.AddressName3 = oBP.Addresses.AddressName3
                                oTargetBP.Addresses.AddressType = oBP.Addresses.AddressType
                                oTargetBP.Addresses.Block = oBP.Addresses.Block
                                oTargetBP.Addresses.City = oBP.Addresses.City
                                oTargetBP.Addresses.County = oBP.Addresses.County
                                oTargetBP.Addresses.Country = oBP.Addresses.Country
                                oTargetBP.Addresses.StreetNo = oBP.Addresses.StreetNo
                                oTargetBP.Addresses.TypeOfAddress = oBP.Addresses.TypeOfAddress
                                oTargetBP.Addresses.State = oBP.Addresses.State
                                oTargetBP.Addresses.ZipCode = oBP.Addresses.ZipCode
                                oTargetBP.Addresses.Street = oBP.Addresses.Street
                                oTargetBP.Addresses.BuildingFloorRoom = oBP.Addresses.BuildingFloorRoom
                                oTargetBP.Addresses.GlobalLocationNumber = oBP.Addresses.GlobalLocationNumber

                                oTargetBP.Addresses.UserFields.Fields.Item("U_FAX").Value = oBP.Addresses.UserFields.Fields.Item("U_FAX").Value
                                oTargetBP.Addresses.UserFields.Fields.Item("U_PIC").Value = oBP.Addresses.UserFields.Fields.Item("U_PIC").Value
                                oTargetBP.Addresses.UserFields.Fields.Item("U_Mobile").Value = oBP.Addresses.UserFields.Fields.Item("U_Mobile").Value
                                oTargetBP.Addresses.UserFields.Fields.Item("U_Email").Value = oBP.Addresses.UserFields.Fields.Item("U_Email").Value
                                oTargetBP.Addresses.UserFields.Fields.Item("U_Phone").Value = oBP.Addresses.UserFields.Fields.Item("U_Phone").Value
                                oTargetBP.Addresses.UserFields.Fields.Item("U_Forwarder").Value = oBP.Addresses.UserFields.Fields.Item("U_Forwarder").Value
                                oTargetBP.Addresses.Add()
                            Next
                            oTargetBP.BilltoDefault = oBP.BilltoDefault
                            oTargetBP.ShipToDefault = oBP.ShipToDefault

                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Payment Code", sFuncName)
                        orsGroup = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("SELECT B.""PymntGroup"" FROM ""OCRD"" A INNER JOIN ""OCTG"" B ON B.""GroupNum"" = A.""GroupNum"" WHERE A.""CardCode"" = '{0}'", oBP.CardCode))
                        If orsGroup.RecordCount = 1 Then
                            sSrcPymntTrmsCod = orsGroup.Fields.Item(0).Value
                        End If
                        If sSrcPymntTrmsCod <> "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                            orsGroup = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orsGroup.DoQuery(String.Format("SELECT ""GroupNum"" FROM ""OCTG"" WHERE ""PymntGroup"" = '{0}'", sSrcPymntTrmsCod))
                            If orsGroup.RecordCount = 1 Then
                                oTargetBP.PayTermsGrpCode = orsGroup.Fields.Item(0).Value
                            End If
                        End If

                        'oTargetBP.PayTermsGrpCode = oBP.PayTermsGrpCode
                        oTargetBP.IntrestRatePercent = oBP.IntrestRatePercent
                        oTargetBP.PriceListNum = oBP.PriceListNum
                        oTargetBP.DiscountPercent = oBP.DiscountPercent

                        oTargetBP.CreditLimit = oBP.CreditLimit
                        oTargetBP.MaxCommitment = oBP.MaxCommitment
                        oTargetBP.EffectiveDiscount = oBP.EffectiveDiscount

                        oTargetBP.HouseBank = oBP.HouseBank
                        oTargetBP.HouseBankAccount = oBP.HouseBankAccount
                        oTargetBP.HouseBankBranch = oBP.HouseBankBranch
                        oTargetBP.HouseBankCountry = oBP.HouseBankCountry
                        oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode

                        For imjs As Integer = 1 To oBP.BPPaymentMethods.PaymentMethodCode.Count
                            oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode
                            oTargetBP.BPPaymentMethods.Add()
                            'oRset_Tar.MoveNext()
                        Next imjs

                        ' ''BP Bank Details 
                        If oTargetBP.BPBankAccounts.Count > 0 Then

                            Dim delete As Boolean = False
                            For i As Integer = 0 To oTargetBP.BPBankAccounts.Count - 1
                                oTargetBP.BPBankAccounts.SetCurrentLine(oTargetBP.BPBankAccounts.Count - 1)
                                oTargetBP.BPBankAccounts.Delete()
                                If oTargetBP.BPBankAccounts.Count = 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        For i As Integer = 0 To oBP.BPBankAccounts.Count - 1
                            oBP.BPBankAccounts.SetCurrentLine(i)
                            'orsTarget.DoQuery(String.Format("SELECT ""BankCode"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.BPBankAccounts.BankCode))
                            'If orsTarget.RecordCount = 1 Then
                            oTargetBP.BPBankAccounts.BankCode = oBP.BPBankAccounts.BankCode
                            oTargetBP.BPBankAccounts.Country = oBP.BPBankAccounts.Country
                            oTargetBP.BPBankAccounts.BPCode = oBP.BPBankAccounts.BPCode
                            oTargetBP.BPBankAccounts.AccountNo = oBP.BPBankAccounts.AccountNo
                            oTargetBP.BPBankAccounts.AccountName = oBP.BPBankAccounts.AccountName
                            oTargetBP.BPBankAccounts.Branch = oBP.BPBankAccounts.Branch
                            oTargetBP.BPBankAccounts.BICSwiftCode = oBP.BPBankAccounts.BICSwiftCode
                            oTargetBP.BPBankAccounts.InternalKey = oBP.BPBankAccounts.InternalKey
                            oTargetBP.BPBankAccounts.ControlKey = oBP.BPBankAccounts.ControlKey
                            oTargetBP.BPBankAccounts.IBAN = oBP.BPBankAccounts.IBAN
                            oTargetBP.BPBankAccounts.Street = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.State = oBP.BPBankAccounts.State
                            oTargetBP.BPBankAccounts.Block = oBP.BPBankAccounts.Block
                            oTargetBP.BPBankAccounts.BuildingFloorRoom = oBP.BPBankAccounts.BuildingFloorRoom
                            oTargetBP.BPBankAccounts.City = oBP.BPBankAccounts.City
                            oTargetBP.BPBankAccounts.MandateID = oBP.BPBankAccounts.MandateID

                            oTargetBP.BPBankAccounts.Add()
                            'Else
                            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Bank code is not available in target bank setup " & oBP.BPBankAccounts.BankCode, sFuncName)
                            'Console.WriteLine("In Target Company Bank Setup the Bank Code- " & oBP.BPBankAccounts.BankCode, sFuncName & "doesn't exist ")
                            'End If

                        Next
                        orsTarget.DoQuery(String.Format("SELECT ""BankCode"",""CountryCod"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.DefaultBankCode))
                        If orsTarget.RecordCount = 1 Then
                            Dim sBankcode As String = orsTarget.Fields.Item(0).Value
                            Dim sCtrycode As String = orsTarget.Fields.Item(1).Value

                            oTargetBP.DefaultBankCode = oBP.DefaultBankCode
                            oTargetBP.DefaultAccount = oBP.DefaultAccount

                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Contact Employees", sFuncName)

                        If oTargetBP.ContactEmployees.Count = 1 Then
                            For imjs As Integer = 0 To oBP.ContactEmployees.Count - 1
                                oBP.ContactEmployees.SetCurrentLine(imjs)
                                'oTargetBP.ContactEmployees.Add()
                                oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                oTargetBP.ContactEmployees.Add()
                            Next
                        ElseIf oTargetBP.ContactEmployees.Count > 0 Then
                            For imjs As Integer = 0 To oBP.ContactEmployees.Count - 1
                                oBP.ContactEmployees.SetCurrentLine(imjs)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oBP.ContactEmployees.Name " & oBP.ContactEmployees.Name, sFuncName)
                                If oBP.ContactEmployees.Name = "" Then Continue For
                                oDVContact.RowFilter = "Name='" & oBP.ContactEmployees.Name & "'"
                                If oDVContact.Count > 0 Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Index " & oDVContact(0)("No").ToString(), sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Name " & oDVContact(0)("Name").ToString(), sFuncName)
                                    oTargetBP.ContactEmployees.SetCurrentLine(Convert.ToInt32(oDVContact(0)("No").ToString()))
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigned", sFuncName)
                                    oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                    oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                    oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                    oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                    oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                    oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                    oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                    oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                    oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                    oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                    oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                    oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                    oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                    oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                    oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                    oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                    oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                    oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                    oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                    oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                    oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                Else

                                    oTargetBP.ContactEmployees.Add()
                                    oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                    oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                    oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                    oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                    oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                    oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                    oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                    oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                    oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                    oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                    oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                    oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                    oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                    oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                    oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                    oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                    oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                    oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                    oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                    oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                    oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                End If
                                ''oTargetBP.ContactEmployees.Add()
                            Next
                        End If
                        oTargetBP.ContactPerson = oBP.ContactPerson
                        oTargetBP.UserFields.Fields.Item("U_PreferredFowarder").Value = oBP.UserFields.Fields.Item("U_PreferredFowarder").Value
                        oTargetBP.UserFields.Fields.Item("U_LeadSource").Value = oBP.UserFields.Fields.Item("U_LeadSource").Value
                        oTargetBP.UserFields.Fields.Item("U_FwdrAddr").Value = oBP.UserFields.Fields.Item("U_FwdrAddr").Value
                        oTargetBP.UserFields.Fields.Item("U_AnnumBudgSales").Value = oBP.UserFields.Fields.Item("U_AnnumBudgSales").Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP " & CardCode, sFuncName)
                        If oTargetBP.Update() <> 0 Then
                            targetCompany.GetLastError(sErrCode, sErrMsg)
                            Throw New Exception("Could not update BP to Target Company" + " - " + sErrMsg)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & "Could not update BP to Target Company" + " - " + sErrMsg, sFuncName)
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function CreateBPMaster()", sFuncName)
                        If CreateBPMaster(CardCode, oTargetBP, oBP, targetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If

                UpdateBPMaster = RTN_SUCCESS
                sErrDesc = String.Empty

            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                UpdateBPMaster = RTN_ERROR
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBP)
                oBP = Nothing
                orsGroup = Nothing
                oTargetBP = Nothing

            End Try
        Else
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error: CardCode not found!!!!", sFuncName)
            UpdateBPMaster = RTN_ERROR
            sErrDesc = "CardCode not found"
        End If
    End Function

    Public Function CreateBPMaster(ByVal CardCode As String, ByVal oTargetBP As SAPbobsCOM.BusinessPartners, ByVal oBP As SAPbobsCOM.BusinessPartners, ByVal targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim ors As SAPbobsCOM.Recordset = Nothing
        Dim orsTarget As SAPbobsCOM.Recordset = Nothing
        Dim GroupName As String = ""
        Dim sSrcShipTypeName As String = String.Empty
        Dim sSrcPymntTrmsCod As String = String.Empty

        Dim sFuncName = "CreateBPMaster()"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
        If oBP.GetByKey(CardCode) Then
            Try
                ors = p_SrcCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orsTarget = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


                If oTargetBP.GetByKey(oBP.CardCode) Then
                    Throw New Exception(String.Format("BP : {0} aldready existed in Branch : {1}", oBP.CardCode, targetCompany.CompanyName))
                End If
                oTargetBP.CardCode = oBP.CardCode
                oTargetBP.CardName = oBP.CardName
                oTargetBP.CardType = oBP.CardType
                oTargetBP.CardForeignName = oBP.CardForeignName

                oTargetBP.CompanyPrivate = oBP.CompanyPrivate
                If oBP.CardType = SAPbobsCOM.BoCardTypes.cCustomer Then
                    oTargetBP.Currency = "##"
                End If
                oTargetBP.DiscountPercent = oBP.DiscountPercent
                oTargetBP.Address = oBP.Address
                oTargetBP.EmailAddress = oBP.EmailAddress
                oTargetBP.Phone1 = oBP.Phone1
                oTargetBP.Phone2 = oBP.Phone2
                oTargetBP.Cellular = oBP.Cellular
                oTargetBP.Fax = oBP.Fax
                oTargetBP.Password = oBP.Password
                oTargetBP.BusinessType = oBP.BusinessType
                oTargetBP.AdditionalID = oBP.AdditionalID
                oTargetBP.VatIDNum = oBP.VatIDNum
                oTargetBP.FederalTaxID = oBP.FederalTaxID
                oTargetBP.Notes = oBP.Notes
                oTargetBP.FreeText = oBP.FreeText
                oTargetBP.AliasName = oBP.AliasName
                oTargetBP.GlobalLocationNumber = oBP.GlobalLocationNumber
                oTargetBP.Valid = oBP.Valid
                oTargetBP.Frozen = oBP.Frozen
                oTargetBP.Website = oBP.Website
                oTargetBP.UnifiedFederalTaxID = oBP.UnifiedFederalTaxID

                ors.DoQuery(String.Format("Select ""GroupName"" from ""OCRG"" where ""GroupCode"" = {0}", oBP.GroupCode))
                GroupName = ors.Fields.Item(0).Value

                ors = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ors.DoQuery(String.Format("Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '{0}'", GroupName))

                If ors.RecordCount = 1 Then
                    oTargetBP.GroupCode = ors.Fields.Item(0).Value
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping Code", sFuncName)
                ors.DoQuery(String.Format("SELECT ""TrnspName"" FROM ""OSHP"" WHERE ""TrnspCode"" = (SELECT ""ShipType"" FROM ""OCRD"" WHERE ""CardCode"" = '{0}')", oBP.CardCode))
                If ors.RecordCount = 1 Then
                    sSrcShipTypeName = ors.Fields.Item(0).Value
                End If
                If sSrcShipTypeName <> "" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                    orsTarget.DoQuery(String.Format("SELECT ""TrnspCode"" FROM ""OSHP"" WHERE ""TrnspName"" = '{0}'", sSrcShipTypeName))
                    If orsTarget.RecordCount = 1 Then
                        oTargetBP.ShippingType = orsTarget.Fields.Item(0).Value
                    End If
                End If

                'oTargetBP.DebitorAccount = oBP.DebitorAccount
                If oTargetBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES Then
                    oTargetBP.ValidFrom = oBP.ValidFrom
                    oTargetBP.ValidTo = oBP.ValidTo
                    oTargetBP.ValidRemarks = oBP.ValidRemarks
                End If
                If oTargetBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
                    oTargetBP.FrozenFrom = oBP.FrozenFrom
                    oTargetBP.FrozenTo = oBP.FrozenTo
                    oTargetBP.FrozenRemarks = oBP.FrozenRemarks
                End If

                oTargetBP.UserFields.Fields.Item("U_PreferredFowarder").Value = oBP.UserFields.Fields.Item("U_PreferredFowarder").Value
                oTargetBP.UserFields.Fields.Item("U_LeadSource").Value = oBP.UserFields.Fields.Item("U_LeadSource").Value
                oTargetBP.UserFields.Fields.Item("U_FwdrAddr").Value = oBP.UserFields.Fields.Item("U_FwdrAddr").Value
                oTargetBP.UserFields.Fields.Item("U_AnnumBudgSales").Value = oBP.UserFields.Fields.Item("U_AnnumBudgSales").Value

                If oBP.Addresses.Count > 0 And oBP.Addresses.AddressName <> "" Then
                    For i As Integer = 0 To oBP.Addresses.Count - 1

                        oBP.Addresses.SetCurrentLine(i)
                        oTargetBP.Addresses.AddressName = oBP.Addresses.AddressName
                        oTargetBP.Addresses.AddressName2 = oBP.Addresses.AddressName2
                        oTargetBP.Addresses.AddressName3 = oBP.Addresses.AddressName3
                        oTargetBP.Addresses.AddressType = oBP.Addresses.AddressType
                        oTargetBP.Addresses.Block = oBP.Addresses.Block
                        oTargetBP.Addresses.City = oBP.Addresses.City
                        oTargetBP.Addresses.County = oBP.Addresses.County
                        oTargetBP.Addresses.Country = oBP.Addresses.Country
                        oTargetBP.Addresses.StreetNo = oBP.Addresses.StreetNo
                        oTargetBP.Addresses.TypeOfAddress = oBP.Addresses.TypeOfAddress
                        oTargetBP.Addresses.State = oBP.Addresses.State
                        oTargetBP.Addresses.ZipCode = oBP.Addresses.ZipCode
                        oTargetBP.Addresses.Street = oBP.Addresses.Street
                        oTargetBP.Addresses.BuildingFloorRoom = oBP.Addresses.BuildingFloorRoom
                        oTargetBP.Addresses.GlobalLocationNumber = oBP.Addresses.GlobalLocationNumber
                        oTargetBP.Addresses.UserFields.Fields.Item("U_FAX").Value = oBP.Addresses.UserFields.Fields.Item("U_FAX").Value
                        oTargetBP.Addresses.UserFields.Fields.Item("U_PIC").Value = oBP.Addresses.UserFields.Fields.Item("U_PIC").Value
                        oTargetBP.Addresses.UserFields.Fields.Item("U_Mobile").Value = oBP.Addresses.UserFields.Fields.Item("U_Mobile").Value
                        oTargetBP.Addresses.UserFields.Fields.Item("U_Email").Value = oBP.Addresses.UserFields.Fields.Item("U_Email").Value
                        oTargetBP.Addresses.UserFields.Fields.Item("U_Phone").Value = oBP.Addresses.UserFields.Fields.Item("U_Phone").Value
                        oTargetBP.Addresses.UserFields.Fields.Item("U_Forwarder").Value = oBP.Addresses.UserFields.Fields.Item("U_Forwarder").Value
                        oTargetBP.Addresses.Add()
                    Next
                    oTargetBP.BilltoDefault = oBP.BilltoDefault
                    oTargetBP.ShipToDefault = oBP.ShipToDefault
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping Code", sFuncName)
                ors.DoQuery(String.Format("SELECT B.""PymntGroup"" FROM ""OCRD"" A INNER JOIN ""OCTG"" B ON B.""GroupNum"" = A.""GroupNum"" WHERE A.""CardCode"" = '{0}'", oBP.CardCode))
                If ors.RecordCount = 1 Then
                    sSrcPymntTrmsCod = ors.Fields.Item(0).Value
                End If
                If sSrcPymntTrmsCod <> "" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                    orsTarget.DoQuery(String.Format("SELECT ""GroupNum"" FROM ""OCTG"" WHERE ""PymntGroup"" = '{0}'", sSrcPymntTrmsCod))
                    If orsTarget.RecordCount = 1 Then
                        oTargetBP.PayTermsGrpCode = orsTarget.Fields.Item(0).Value
                    End If
                End If
                '''''''''''oTargetBP.PayTermsGrpCode = oBP.PayTermsGrpCode
                oTargetBP.IntrestRatePercent = oBP.IntrestRatePercent
                oTargetBP.PriceListNum = oBP.PriceListNum
                oTargetBP.DiscountPercent = oBP.DiscountPercent

                oTargetBP.CreditLimit = oBP.CreditLimit
                oTargetBP.MaxCommitment = oBP.MaxCommitment
                oTargetBP.EffectiveDiscount = oBP.EffectiveDiscount

                oTargetBP.HouseBank = oBP.HouseBank
                oTargetBP.HouseBankAccount = oBP.HouseBankAccount
                oTargetBP.HouseBankBranch = oBP.HouseBankBranch
                oTargetBP.HouseBankCountry = oBP.HouseBankCountry
                oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode

                ''BP Bank Details 
                For i As Integer = 0 To oBP.BPBankAccounts.Count - 1
                    oBP.BPBankAccounts.SetCurrentLine(i)
                    'orsTarget.DoQuery(String.Format("SELECT ""BankCode"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.BPBankAccounts.BankCode))
                    'If orsTarget.RecordCount = 1 Then
                    oTargetBP.BPBankAccounts.BankCode = oBP.BPBankAccounts.BankCode
                    oTargetBP.BPBankAccounts.Country = oBP.BPBankAccounts.Country
                    oTargetBP.BPBankAccounts.BPCode = oBP.BPBankAccounts.BPCode
                    oTargetBP.BPBankAccounts.AccountNo = oBP.BPBankAccounts.AccountNo
                    oTargetBP.BPBankAccounts.AccountName = oBP.BPBankAccounts.AccountName
                    oTargetBP.BPBankAccounts.Branch = oBP.BPBankAccounts.Branch
                    oTargetBP.BPBankAccounts.BICSwiftCode = oBP.BPBankAccounts.BICSwiftCode
                    oTargetBP.BPBankAccounts.InternalKey = oBP.BPBankAccounts.InternalKey
                    oTargetBP.BPBankAccounts.ControlKey = oBP.BPBankAccounts.ControlKey
                    oTargetBP.BPBankAccounts.IBAN = oBP.BPBankAccounts.IBAN
                    oTargetBP.BPBankAccounts.Street = oBP.BPBankAccounts.Street
                    oTargetBP.BPBankAccounts.State = oBP.BPBankAccounts.Street
                    oTargetBP.BPBankAccounts.Block = oBP.BPBankAccounts.Street
                    oTargetBP.BPBankAccounts.BuildingFloorRoom = oBP.BPBankAccounts.BuildingFloorRoom
                    oTargetBP.BPBankAccounts.City = oBP.BPBankAccounts.City
                    oTargetBP.BPBankAccounts.MandateID = oBP.BPBankAccounts.MandateID

                    oTargetBP.BPBankAccounts.Add()
                    'Else
                    'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Bank code is not available in target bank setup " & oBP.BPBankAccounts.BankCode, sFuncName)
                    'Console.WriteLine("In Target Company Bank Setup the Bank Code- " & oBP.BPBankAccounts.BankCode, sFuncName & "doesn't exist ")
                    'End If

                Next

                orsTarget.DoQuery(String.Format("SELECT ""BankCode"",""CountryCod"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.DefaultBankCode))
                If orsTarget.RecordCount = 1 Then
                    Dim sBankcode As String = orsTarget.Fields.Item(0).Value
                    Dim sCtrycode As String = orsTarget.Fields.Item(1).Value

                    oTargetBP.DefaultBankCode = oBP.DefaultBankCode
                    oTargetBP.DefaultAccount = oBP.DefaultAccount

                End If
                For i As Integer = 0 To oBP.ContactEmployees.Count - 1
                    oBP.ContactEmployees.SetCurrentLine(i)
                    oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                    oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                    oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                    oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                    oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                    oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                    oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                    oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                    oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2
                    oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                    oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                    oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                    oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                    oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                    oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.InternalCode
                    oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                    oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                    oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                    oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                    oTargetBP.ContactEmployees.Active = oBP.ContactEmployees.Active
                    oTargetBP.ContactEmployees.Add()
                Next

                oTargetBP.ContactPerson = oBP.ContactPerson

                If oTargetBP.Add() <> 0 Then
                    Dim errCode As Integer
                    Dim errMess As String = ""
                    targetCompany.GetLastError(errCode, errMess)
                    Throw New Exception("Could not create BP to Target Company" + " - " + errMess)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & "Could not create BP to Target Company" + " - " + errMess, sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                End If

                CreateBPMaster = RTN_SUCCESS
                sErrDesc = String.Empty

            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                CreateBPMaster = RTN_ERROR
            End Try
        Else
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error: CardCode not found!!!!", sFuncName)
        End If
    End Function
#End Region

#Region "BP Price list"
    Public Function UpdateBPPriceList(ByVal CardCode As String, ByVal ItemCode As String, ByRef oCompany As SAPbobsCOM.Company, ByRef targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim oBPPriceList As SAPbobsCOM.SpecialPrices = Nothing
        Dim oTargetBPPriceList As SAPbobsCOM.SpecialPrices = Nothing

        Dim sFuncName As String = String.Empty
        sFuncName = "UpdateBPPriceList()"
        Dim sSQL As String = String.Empty

        oBPPriceList = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices)
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
        'Dim oSpecialPricesDataAreas As SAPbobsCOM.SpecialPricesDataAreas = Nothing
        'Dim oTSpecialPricesDataAreas As SAPbobsCOM.SpecialPricesDataAreas = Nothing
        ''JST 1.25-6 VGL_MANSI-USD
        'ItemCode = "JST 1.25-6"
        'CardCode = "VGL_MANSI-USD"

        If oBPPriceList.GetByKey(ItemCode, CardCode) Then
            Dim sErrMsg As String = ""
            Dim sErrCode As Integer = 0
            Try
                If targetCompany.Connected Then
                    Dim bfound As Boolean = False
                    oTargetBPPriceList = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Conected with target Company " & targetCompany.CompanyDB, sFuncName)
                End If

                If oTargetBPPriceList.GetByKey(ItemCode, CardCode) Then
                    'oTargetBPPriceList.CardCode = oBPPriceList.CardCode
                    oTargetBPPriceList.ItemCode = oBPPriceList.ItemCode
                    oTargetBPPriceList.PriceListNum = oBPPriceList.PriceListNum
                    oTargetBPPriceList.DiscountPercent = oBPPriceList.DiscountPercent
                    oTargetBPPriceList.Currency = oBPPriceList.Currency
                    oTargetBPPriceList.Price = oBPPriceList.Price
                    oTargetBPPriceList.SourcePrice = oBPPriceList.SourcePrice

                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQ").Value = oBPPriceList.UserFields.Fields.Item("U_MoQ").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_SupplierLeadTime").Value = oBPPriceList.UserFields.Fields.Item("U_SupplierLeadTime").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_POSEndUser").Value = oBPPriceList.UserFields.Fields.Item("U_POSEndUser").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_CarModel").Value = oBPPriceList.UserFields.Fields.Item("U_CarModel").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_POSApplication").Value = oBPPriceList.UserFields.Fields.Item("U_POSApplication").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_UsageQty").Value = oBPPriceList.UserFields.Fields.Item("U_UsageQty").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_SOPDateMassProd").Value = oBPPriceList.UserFields.Fields.Item("U_SOPDateMassProd").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_Remarks").Value = oBPPriceList.UserFields.Fields.Item("U_Remarks").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_Selected").Value = oBPPriceList.UserFields.Fields.Item("U_Selected").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_COO").Value = oBPPriceList.UserFields.Fields.Item("U_COO").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_ShippingPoint").Value = oBPPriceList.UserFields.Fields.Item("U_ShippingPoint").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_CostLC").Value = oBPPriceList.UserFields.Fields.Item("U_CostLC").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2SPrice").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2SPrice").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2CostLC").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2CostLC").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3SPrice").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3SPrice").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3CostLC").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3CostLC").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_LeadTime_SG").Value = oBPPriceList.UserFields.Fields.Item("U_LeadTime_SG").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_LeadTime_HK").Value = oBPPriceList.UserFields.Fields.Item("U_LeadTime_HK").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_SupplierSPQ").Value = oBPPriceList.UserFields.Fields.Item("U_SupplierSPQ").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQL2QtDt").Value = oBPPriceList.UserFields.Fields.Item("U_MoQL2QtDt").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2QtNo").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2QtNo").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_SupplierQuote").Value = oBPPriceList.UserFields.Fields.Item("U_SupplierQuote").Value

                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtDt").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtDt").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtNo").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtNo").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_QuoteDate").Value = oBPPriceList.UserFields.Fields.Item("U_QuoteDate").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_ValidToDate").Value = oBPPriceList.UserFields.Fields.Item("U_ValidToDate").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_ActualCustomer").Value = oBPPriceList.UserFields.Fields.Item("U_ActualCustomer").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_TargetPrice").Value = oBPPriceList.UserFields.Fields.Item("U_TargetPrice").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_POSCustomer").Value = oBPPriceList.UserFields.Fields.Item("U_POSCustomer").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_OEM").Value = oBPPriceList.UserFields.Fields.Item("U_OEM").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_CartonQty").Value = oBPPriceList.UserFields.Fields.Item("U_CartonQty").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_OLDPRICE").Value = oBPPriceList.UserFields.Fields.Item("U_OLDPRICE").Value
                    oTargetBPPriceList.UserFields.Fields.Item("U_OLDPRICE_DT").Value = oBPPriceList.UserFields.Fields.Item("U_OLDPRICE_DT").Value

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP Price list " & CardCode, sFuncName)
                    If oTargetBPPriceList.Update() <> 0 Then
                        targetCompany.GetLastError(sErrCode, sErrMsg)
                        Throw New Exception("Could not update BP Price List to Target Company" + " - " + sErrMsg)

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & "Could not update BP Price List to Target Company" + " - " + sErrMsg, sFuncName)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

                    End If
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function CreateBPPriceList()", sFuncName)
                    If CreateBPPriceList(CardCode, ItemCode, oTargetBPPriceList, oBPPriceList, targetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
                UpdateBPPriceList = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                UpdateBPPriceList = RTN_ERROR
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBPPriceList)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBPPriceList)
                oBPPriceList = Nothing
                oTargetBPPriceList = Nothing
            End Try
        Else
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error: BP Price List not found!!!!", sFuncName)
        End If

    End Function

    Public Function CreateBPPriceList(ByVal CardCode As String, ByVal ItemCode As String, ByVal oTargetBPPriceList As SAPbobsCOM.SpecialPrices, ByVal oBPPriceList As SAPbobsCOM.SpecialPrices, ByVal targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName = "CreateBPPriceList()"
        'Dim oSpecialPricesDataAreas As SAPbobsCOM.SpecialPricesDataAreas = Nothing
        'Dim oTSpecialPricesDataAreas As SAPbobsCOM.SpecialPricesDataAreas = Nothing

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
        If oBPPriceList.GetByKey(ItemCode, CardCode) Then
            Try
                If oTargetBPPriceList.GetByKey(oBPPriceList.ItemCode, oBPPriceList.CardCode) Then
                    Throw New Exception(String.Format("CardCode : {0} already existed in Target Company : {1}", oBPPriceList.CardCode, targetCompany.CompanyName))
                End If
                oTargetBPPriceList.CardCode = oBPPriceList.CardCode
                oTargetBPPriceList.ItemCode = oBPPriceList.ItemCode
                oTargetBPPriceList.PriceListNum = oBPPriceList.PriceListNum
                oTargetBPPriceList.DiscountPercent = oBPPriceList.DiscountPercent
                oTargetBPPriceList.Currency = oBPPriceList.Currency
                oTargetBPPriceList.Price = oBPPriceList.Price                
                oTargetBPPriceList.SourcePrice = oBPPriceList.SourcePrice

                oTargetBPPriceList.UserFields.Fields.Item("U_MoQ").Value = oBPPriceList.UserFields.Fields.Item("U_MoQ").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_SupplierLeadTime").Value = oBPPriceList.UserFields.Fields.Item("U_SupplierLeadTime").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_POSEndUser").Value = oBPPriceList.UserFields.Fields.Item("U_POSEndUser").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_CarModel").Value = oBPPriceList.UserFields.Fields.Item("U_CarModel").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_POSApplication").Value = oBPPriceList.UserFields.Fields.Item("U_POSApplication").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_UsageQty").Value = oBPPriceList.UserFields.Fields.Item("U_UsageQty").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_SOPDateMassProd").Value = oBPPriceList.UserFields.Fields.Item("U_SOPDateMassProd").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_Remarks").Value = oBPPriceList.UserFields.Fields.Item("U_Remarks").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_Selected").Value = oBPPriceList.UserFields.Fields.Item("U_Selected").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_COO").Value = oBPPriceList.UserFields.Fields.Item("U_COO").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_ShippingPoint").Value = oBPPriceList.UserFields.Fields.Item("U_ShippingPoint").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_CostLC").Value = oBPPriceList.UserFields.Fields.Item("U_CostLC").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2SPrice").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2SPrice").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2CostLC").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2CostLC").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3SPrice").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3SPrice").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3CostLC").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3CostLC").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_LeadTime_SG").Value = oBPPriceList.UserFields.Fields.Item("U_LeadTime_SG").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_LeadTime_HK").Value = oBPPriceList.UserFields.Fields.Item("U_LeadTime_HK").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_SupplierSPQ").Value = oBPPriceList.UserFields.Fields.Item("U_SupplierSPQ").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQL2QtDt").Value = oBPPriceList.UserFields.Fields.Item("U_MoQL2QtDt").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl2QtNo").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl2QtNo").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_SupplierQuote").Value = oBPPriceList.UserFields.Fields.Item("U_SupplierQuote").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtDt").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtDt").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtNo").Value = oBPPriceList.UserFields.Fields.Item("U_MoQLvl3QtNo").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_QuoteDate").Value = oBPPriceList.UserFields.Fields.Item("U_QuoteDate").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_ValidToDate").Value = oBPPriceList.UserFields.Fields.Item("U_ValidToDate").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_ActualCustomer").Value = oBPPriceList.UserFields.Fields.Item("U_ActualCustomer").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_TargetPrice").Value = oBPPriceList.UserFields.Fields.Item("U_TargetPrice").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_POSCustomer").Value = oBPPriceList.UserFields.Fields.Item("U_POSCustomer").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_OEM").Value = oBPPriceList.UserFields.Fields.Item("U_OEM").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_CartonQty").Value = oBPPriceList.UserFields.Fields.Item("U_CartonQty").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_OLDPRICE").Value = oBPPriceList.UserFields.Fields.Item("U_OLDPRICE").Value
                oTargetBPPriceList.UserFields.Fields.Item("U_OLDPRICE_DT").Value = oBPPriceList.UserFields.Fields.Item("U_OLDPRICE_DT").Value


                If oTargetBPPriceList.Add() <> 0 Then
                    Dim errCode As Integer
                    Dim errMess As String = ""
                    targetCompany.GetLastError(errCode, errMess)
                    Throw New Exception("Could not create BP Price list to Target Company- " & CardCode + " - " & ItemCode + " - " + errMess)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & "Could not create BP Price list to Target Company- " & CardCode + " - " & ItemCode + " - " + errMess, sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                End If

                CreateBPPriceList = RTN_SUCCESS
                sErrDesc = String.Empty

            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                CreateBPPriceList = RTN_ERROR

            End Try
        End If
    End Function
#End Region

#Region "Common Function"
    Public Function ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset, ByRef sErrDesc As String) As DataTable

        Dim NewCol As DataColumn
        Dim NewRow As DataRow
        Dim ColCount As Integer
        Dim dtTable As DataTable = Nothing


        Try
            dtTable = New DataTable

            For ColCount = 0 To SAPRecordset.Fields.Count - 1
                NewCol = New DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                dtTable.Columns.Add(NewCol)
            Next

            Do Until SAPRecordset.EoF
                NewRow = dtTable.NewRow
                'populate each column in the row we're creating
                For ColCount = 0 To SAPRecordset.Fields.Count - 1
                    NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value
                Next

                'Add the row to the datatable
                dtTable.Rows.Add(NewRow)
                SAPRecordset.MoveNext()
            Loop
            Return dtTable
        Catch ex As Exception
            sErrDesc = ex.Message
            Exit Function
        End Try


    End Function
#End Region

End Module


