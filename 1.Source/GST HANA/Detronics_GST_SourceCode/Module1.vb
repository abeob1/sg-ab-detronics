Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports System.Threading.Thread
Imports System.Threading
Module Module1
    Public oGeneralService As SAPbobsCOM.GeneralService
    Public oGeneralData As SAPbobsCOM.GeneralData
    Public oSons As SAPbobsCOM.GeneralDataCollection
    Public oSon As SAPbobsCOM.GeneralData
    Public sCmp As SAPbobsCOM.CompanyService
    Public oForm As SAPbouiCOM.Form
    Public oForm1 As SAPbouiCOM.Form
    Public oRecordSet As SAPbobsCOM.Recordset
    Public oRecordSet1 As SAPbobsCOM.Recordset
    Public oRecordSet2 As SAPbobsCOM.Recordset
    Public oRecordSet3 As SAPbobsCOM.Recordset
    Public format1 As New System.Globalization.CultureInfo("fr-FR", True)
    Public DocEtry As Integer
  

    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Dim oXmlDoc As New Xml.XmlDocument
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        oXmlDoc.Load(sPath & "\GK_FM\" & FileName)
        Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub
  


End Module
