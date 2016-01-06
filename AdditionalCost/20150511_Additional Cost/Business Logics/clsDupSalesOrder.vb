Public Class clsDupSalesOrder
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_DupSales, frm_DupSales)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.DataSources.UserDataSources.Add("Interval", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
        oApplication.Utilities.setUserDatabind(oForm, "41", "Interval")
        oForm.DataSources.UserDataSources.Add("SalesNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "1000007", "SalesNo")
        oForm.DataSources.UserDataSources.Add("SalDocNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "32", "SalDocNum")
        oEditText = oForm.Items.Item("1000007").Specific
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "DocEntry"
        oForm.DataSources.UserDataSources.Add("DupSale", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
        oApplication.Utilities.setUserDatabind(oForm, "1000002", "DupSale")
        oForm.DataSources.UserDataSources.Add("StDate", SAPbouiCOM.BoDataType.dt_DATE)
        oApplication.Utilities.setUserDatabind(oForm, "18", "StDate")
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub


    Public Sub LoadForm(ByVal aDocNum As String)
        oForm = oApplication.Utilities.LoadForm(xml_DupSales, frm_DupSales)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.DataSources.UserDataSources.Add("Interval", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
        oApplication.Utilities.setUserDatabind(oForm, "41", "Interval")
        oForm.DataSources.UserDataSources.Add("SalesNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "1000007", "SalesNo")
        oForm.DataSources.UserDataSources.Add("SalDocNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "32", "SalDocNum")
        oEditText = oForm.Items.Item("1000007").Specific
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "DocEntry"
        oForm.DataSources.UserDataSources.Add("DupSale", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
        oApplication.Utilities.setUserDatabind(oForm, "1000002", "DupSale")
        oForm.DataSources.UserDataSources.Add("StDate", SAPbouiCOM.BoDataType.dt_DATE)
        oApplication.Utilities.setUserDatabind(oForm, "18", "StDate")
        Dim oRs As SAPbobsCOM.Recordset
        oRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery("Select DocEntry from ORDR where DocNum=" & aDocNum)
        oApplication.Utilities.SetEditText(oForm, "1000007", oRs.Fields.Item(0).Value)
        oApplication.Utilities.SetEditText(oForm, "32", aDocNum)
        oForm.PaneLevel = 2
        oForm.Freeze(False)
    End Sub
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Gridbind(ByVal DocNum As Integer)
        Dim strqry As String
        oGrid = oForm.Items.Item("10").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_0")
        strqry = "SELECT T0.DocEntry,T1.DocNum,ItemCode,Quantity,Price,LineTotal,U_Z_Winch FROM RDR1 T0 inner Join ORDR T1 on T1.DocEntry=T0.DocEntry Where T1.DocEntry = " & DocNum
        oGrid.DataTable.ExecuteQuery(strqry)
        oGrid.Columns.Item("DocNum").TitleObject.Caption = "Document Number"
        oGrid.Columns.Item("DocNum").Editable = False
        oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Document Entry"
        oGrid.Columns.Item("DocEntry").Editable = False
        oEditTextColumn = oGrid.Columns.Item("DocEntry")
        oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order
        oGrid.Columns.Item("ItemCode").TitleObject.Caption = "Item Code"
        oGrid.Columns.Item("ItemCode").Editable = False
        oGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
        oGrid.Columns.Item("Quantity").Editable = False
        oGrid.Columns.Item("Price").TitleObject.Caption = "UnitPrice"
        oGrid.Columns.Item("Price").Editable = False
        oGrid.Columns.Item("LineTotal").TitleObject.Caption = "Total Uniit Price"
        oGrid.Columns.Item("LineTotal").Editable = False
        oGrid.Columns.Item("U_Z_Winch").TitleObject.Caption = "Winch Plate No"
        oGrid.Columns.Item("U_Z_Winch").Editable = False
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
    Private Sub Gridbind1(ByVal DocNum As Integer)
        Dim strqry As String
        oGrid = oForm.Items.Item("25").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_1")
        strqry = "SELECT U_Z_SourceDoc,U_Z_SourceDocNum,U_Z_DupEntry,U_Z_StartingDate,DateName(dw,U_Z_StartingDate) 'Day',U_Z_Winch FROM [@DORDR] Where U_Z_SourceDoc = " & DocNum
        oGrid.DataTable.ExecuteQuery(strqry)
        oGrid.Columns.Item("U_Z_SourceDocNum").TitleObject.Caption = "Document Number"
        oGrid.Columns.Item("U_Z_SourceDocNum").Editable = False
        oGrid.Columns.Item("U_Z_SourceDoc").TitleObject.Caption = "Document Number"
        oGrid.Columns.Item("U_Z_SourceDoc").Editable = False
        oEditTextColumn = oGrid.Columns.Item("U_Z_SourceDoc")
        oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order
        oGrid.Columns.Item("U_Z_DupEntry").TitleObject.Caption = "Duplicate Entry"
        oGrid.Columns.Item("U_Z_DupEntry").Editable = False
        oGrid.Columns.Item("U_Z_StartingDate").TitleObject.Caption = "Delivery Date"
        oGrid.Columns.Item("U_Z_StartingDate").Editable = True
        oGrid.Columns.Item("Day").TitleObject.Caption = "Day"
        oGrid.Columns.Item("Day").Editable = False
        oGrid.Columns.Item("U_Z_Winch").TitleObject.Caption = "Winch Plate No"
        oGrid.Columns.Item("U_Z_Winch").Editable = True
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next
    End Sub
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim SalesNo, Reqno, posname As String
            Reqno = oApplication.Utilities.getEdittextvalue(aForm, "18")
            posname = oApplication.Utilities.getEdittextvalue(aForm, "1000002")
            SalesNo = oApplication.Utilities.getEdittextvalue(aForm, "1000007")
            If oForm.PaneLevel = 2 Then
                If SalesNo = "" Then
                    oApplication.Utilities.Message("Sales Order No is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            If oForm.PaneLevel = 3 Then
                If Reqno = "" Then
                    oApplication.Utilities.Message("Number of duplicate entry is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If posname = "" Then
                    oApplication.Utilities.Message("Delivery Starting Date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
          
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function Validation1(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            oGrid = oForm.Items.Item("25").Specific
            For intX As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                Dim delDate As String = oGrid.DataTable.GetValue("U_Z_StartingDate", intX)
                If delDate = Nothing Then
                    oApplication.Utilities.Message("Delivery Date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#Region "AddToUDT"
    Private Function AddToUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strTable, strEmpId, strCode As String
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Delete from [@DORDR] where U_Z_SourceDoc='" & oApplication.Utilities.getEdittextvalue(aForm, "1000007") & "'")
        oUserTable = oApplication.Company.UserTables.Item("DORDR")
        oGrid = aForm.Items.Item("10").Specific
        Dim dupEntry, dateincrement As Integer
        Dim Interval As Double
        Dim StartDupdt As Date
        dupEntry = CInt(oApplication.Utilities.getEdittextvalue(aForm, "1000002"))
        Interval = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "41"))
        StartDupdt = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "18"))
        strTable = "@DORDR"
        For DupRow As Integer = 0 To dupEntry - 1
            strCode = oApplication.Utilities.getMaxCode(strTable, "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_DupEntry").Value = DupRow + 1
            oUserTable.UserFields.Fields.Item("U_Z_SourceDoc").Value = oApplication.Utilities.getEdittextvalue(aForm, "1000007")
            oUserTable.UserFields.Fields.Item("U_Z_SourceDocNum").Value = oApplication.Utilities.getEdittextvalue(aForm, "32")
            oRec.DoQuery("Select isnull(U_Z_Winch,0) from ORDR where DocEntry=" & oApplication.Utilities.getEdittextvalue(aForm, "1000007"))
            oUserTable.UserFields.Fields.Item("U_Z_Winch").Value = oRec.Fields.Item(0).Value
            If Interval > 0 Then
                dateincrement = Interval
            Else
                dateincrement = DupRow
            End If
            oUserTable.UserFields.Fields.Item("U_Z_StartingDate").Value = StartDupdt '.AddDays(dateincrement)

            If dateincrement <= 0 Then
                dateincrement = 1
            End If
            StartDupdt = StartDupdt.AddDays(dateincrement)
            If oUserTable.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        Next
        Return True
    End Function
    Private Function CrateDuplicateSalesOrderXML() As Boolean
        Dim pDraft As SAPbobsCOM.Documents
        pDraft = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        Dim pOrder As SAPbobsCOM.Documents
        pOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        oApplication.Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_AllNodes
        oApplication.Company.XMLAsString = False

        oGrid = oForm.Items.Item("25").Specific
        For intX As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            pDraft.GetByKey(oGrid.DataTable.GetValue("U_Z_SourceDoc", intX))
            pDraft.SaveXML("c:\drafts.xml")
            pOrder = oApplication.Company.GetBusinessObjectFromXML("c:\drafts.xml", 0)
            If pOrder.Add() <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Next
      
    End Function
    Private Function CrateDuplicateSalesOrder() As Boolean
        If oApplication.Company.InTransaction() Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim dupEntry, DocNum As Integer
        Dim StartDupdt As Date
        dupEntry = CInt(oApplication.Utilities.getEdittextvalue(oForm, "1000002"))
        DocNum = CInt(oApplication.Utilities.getEdittextvalue(oForm, "1000007"))
        StartDupdt = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(oForm, "18"))
        oGrid = oForm.Items.Item("25").Specific
        For intX As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If SalesOrer_Duplicate(oGrid.DataTable.GetValue("U_Z_SourceDoc", intX), oGrid.DataTable.GetValue("U_Z_StartingDate", intX), oGrid.DataTable.GetValue("U_Z_Winch", intX)) = False Then
                If oApplication.Company.InTransaction() Then
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
                Return False
            End If
        Next
        oRec.DoQuery("Update ORDR set U_Z_DupEntry=" & dupEntry & ",U_Z_StartingDate='" & StartDupdt.ToString("yyyy-MM-dd") & "' where DocNum=" & DocNum)
        If oApplication.Company.InTransaction() Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
    Private Function SalesOrer_Duplicate(ByVal aDocEntry As Integer, ByVal aDocdate As Date, ByVal aWinch As Integer) As Boolean
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim objremoteRec As SAPbobsCOM.Recordset
        objremoteRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry from ORDR where DocEntry=" & aDocEntry)
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objremoteDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                objMainDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                objMainDoc.Address = objremoteDoc.Address
                objMainDoc.Address2 = objremoteDoc.Address2
                objMainDoc.Series = objremoteDoc.Series
                objMainDoc.CardCode = objremoteDoc.CardCode
                objMainDoc.CardName = objremoteDoc.CardName
                objMainDoc.CentralBankIndicator = objremoteDoc.CentralBankIndicator
                objMainDoc.ClosingRemarks = objremoteDoc.ClosingRemarks
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                objMainDoc.DeferredTax = objremoteDoc.DeferredTax
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.DocDate = objremoteDoc.DocDate
                objMainDoc.DocDueDate = aDocdate ' objremoteDoc.DocDueDate
                ' objMainDoc.DocRate = objremoteDoc.DocRate
                objMainDoc.DocTotal = objremoteDoc.DocTotal
                objMainDoc.DocType = objremoteDoc.DocType
                objMainDoc.DocumentSubType = objremoteDoc.DocumentSubType
                objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                objMainDoc.TaxDate = objremoteDoc.TaxDate
                objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                ' objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                objMainDoc.UserFields.Fields.Item("U_Z_SourceDoc").Value = aDocEntry

                If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                Else
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                objMainDoc.UserFields.Fields.Item("U_Z_PumpCode").Value = objremoteDoc.UserFields.Fields.Item("U_Z_PumpCode").Value

                'objMainDoc.UserFields.Fields.Item(intLoop).Value = objremoteDoc.UserFields.Fields.Item(intLoop)

                For intLoop As Integer = 0 To objremoteDoc.UserFields.Fields.Count - 1
                    Try
                        objMainDoc.UserFields.Fields.Item(intLoop).Value = objremoteDoc.UserFields.Fields.Item(intLoop).Value
                    Catch ex As Exception
                    End Try

                Next
                objMainDoc.UserFields.Fields.Item("U_Z_Winch").Value = aWinch
                ' objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                    If objremoteDoc.Expenses.LineTotal > 0 Then
                        If IntExp > 0 Then
                            objMainDoc.Expenses.Add()
                            objMainDoc.Expenses.SetCurrentLine(IntExp)
                        End If
                        objremoteDoc.Expenses.SetCurrentLine(IntExp)
                        objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                        objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                        objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                        objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                        objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                        objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                        objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                        objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                        objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                        objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                        objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                    End If
                Next


                For intLoop As Integer = 0 To objremoteDoc.UserFields.Fields.Count - 1
                    Try
                        objMainDoc.Lines.UserFields.Fields.Item(intLoop).Value = objremoteDoc.Lines.UserFields.Fields.Item(intLoop).Value
                    Catch ex As Exception
                    End Try

                Next
                For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                    If intLoop > 0 Then
                        objMainDoc.Lines.Add()
                        objMainDoc.Lines.SetCurrentLine(intLoop)
                    End If
                    objremoteDoc.Lines.SetCurrentLine(intLoop)
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                Next
                If objMainDoc.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
        Next
        Return True
    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_DupSales Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 2 Or oForm.PaneLevel = 3) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "22" Then
                                    If Validation1(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.Freeze(True)
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        If oForm.PaneLevel = 3 Then
                                            Dim IntDocNum As Integer = oApplication.Utilities.getEdittextvalue(oForm, "1000007")
                                            Gridbind(IntDocNum)
                                        ElseIf oForm.PaneLevel = 4 Then
                                            If AddToUDT(oForm) = True Then
                                                Dim IntDocNum As Integer = oApplication.Utilities.getEdittextvalue(oForm, "1000007")
                                                Gridbind1(IntDocNum)
                                            End If
                                        End If
                                        oForm.Freeze(False)
                                    Case "4"
                                        oForm.Freeze(True)
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        oForm.Freeze(False)
                                    Case "22"
                                        If oApplication.SBO_Application.MessageBox("Do you want confirm the Duplicate Sales Order", , "Yes", "No") = 2 Then
                                            Exit Sub
                                        ElseIf CrateDuplicateSalesOrder() = True Then
                                            oApplication.Utilities.Message("Duplicate Sales Order Created successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            oForm.Close()
                                        Else
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1, val As String
                                Dim sCHFL_ID As String
                                Dim intChoice As Integer


                                oCFLEvento = pVal
                                sCHFL_ID = oCFLEvento.ChooseFromListUID
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                If (oCFLEvento.BeforeAction = False) Then
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects
                                    intChoice = 0
                                    oForm.Freeze(True)

                                    If pVal.ItemUID = "1000007" Then
                                        val1 = oDataTable.GetValue("DocNum", 0)
                                        val = oDataTable.GetValue("DocEntry", 0)
                                        Try
                                            oApplication.Utilities.setEdittextvalue(oForm, "32", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "1000007", val)
                                        Catch ex As Exception
                                            oForm.Freeze(False)
                                        End Try
                                    End If

                                    oForm.Freeze(False)
                                End If
                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_DupSales
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
