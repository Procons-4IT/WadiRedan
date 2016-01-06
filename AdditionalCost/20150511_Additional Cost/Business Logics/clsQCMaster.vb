
Public Class clsQCMaster
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oCheckbox, oCheckbox1 As SAPbouiCOM.CheckBox
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private oTemp As SAPbobsCOM.Recordset
    Private InvBaseDocNo, strname As String
    Private InvForConsumedItems As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
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
            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_ItemType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub LoadForm(ByVal objSourceForm As SAPbouiCOM.Form, ByVal Rowid As String, ByVal RefCode As String)
        oForm = oApplication.Utilities.LoadForm(xml_AddCostDetails, frm_AddCostDetails)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Dim oRec As SAPbobsCOM.Recordset
        Dim strqry As String
        oCheckbox = oForm.Items.Item("10").Specific
        oCheckbox1 = oForm.Items.Item("12").Specific
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strlinestatus As String
        oRec.DoQuery("Select * from RDR1 where U_Z_AddRef='" & RefCode & "'")
        If oRec.RecordCount > 0 Then
            If oRec.Fields.Item("Quantity").Value <> oRec.Fields.Item("OpenQty").Value Then
                strlinestatus = "C"
            Else
                strlinestatus = oRec.Fields.Item("LineStatus").Value
            End If

        Else
            strlinestatus = "O"
        End If
        oCombobox = oForm.Items.Item("26").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("O", "Open")
        oCombobox.ValidValues.Add("C", "Close")
        oCombobox.Select(strlinestatus, SAPbouiCOM.BoSearchKey.psk_ByValue)
        oForm.Items.Item("26").DisplayDesc = True

        strqry = "Select * from [@Z_ADC1] where Code='" & RefCode & "'"
        oRec.DoQuery(strqry)
        If oRec.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(oForm, "18", oRec.Fields.Item("U_Z_WhPumpPrice").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "19", oRec.Fields.Item("U_Z_WatProofPrice").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "24", oRec.Fields.Item("U_Z_TempValue").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "4", oRec.Fields.Item("U_Z_RefCode").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "22", oRec.Fields.Item("U_Z_DefTemp").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "17", Rowid)
            oApplication.Utilities.setEdittextvalue(oForm, "6", oRec.Fields.Item("U_Z_ItemCode").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "8", oRec.Fields.Item("U_Z_UnitPrice").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "14", oRec.Fields.Item("U_Z_Temp").Value)
            Dim dblprice As Double = oRec.Fields.Item("U_Z_ActualPrice").Value

            If oRec.Fields.Item("U_Z_WhoutPump").Value = "Y" Then
                oCheckbox.Checked = True
            Else
                oCheckbox.Checked = False
            End If
            If oRec.Fields.Item("U_Z_WaterProof").Value = "Y" Then
                oCheckbox1.Checked = True
            Else
                oCheckbox1.Checked = False
            End If
            oApplication.Utilities.setEdittextvalue(oForm, "16", dblprice) ' oRec.Fields.Item("U_Z_ActualPrice").Value)
        End If
        oGrid = oForm.Items.Item("27").Specific
        oGrid.DataTable.ExecuteQuery("Select * from [@Z_ADC2] where U_Z_AddRef='" & RefCode & "'")
        formatgrid(oForm)

        oForm.Items.Item("124").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        If strlinestatus = "C" Then
            oForm.Items.Item("10").Enabled = False
            oForm.Items.Item("12").Enabled = False
            oForm.Items.Item("14").Enabled = False
            oForm.Items.Item("3").Visible = False
        Else
            oForm.Items.Item("10").Enabled = True
            oForm.Items.Item("12").Enabled = True
            oForm.Items.Item("14").Enabled = True
            oForm.Items.Item("3").Visible = True
        End If
        oForm.Freeze(False)
    End Sub

    Private Sub addrow(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        oGrid = aForm.Items.Item("27").Specific
        If oGrid.DataTable.Rows.Count - 1 < 0 Then
            oGrid.DataTable.Rows.Add()
        Else
            If oGrid.DataTable.GetValue("U_Z_ItemCode", oGrid.DataTable.Rows.Count - 1) <> "" Then
                oGrid.DataTable.Rows.Add()
            End If
        End If
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intRow, intRow + 1)

        Next
        aForm.Freeze(False)
    End Sub

    Private Sub commintTrans(ByVal aChoice As String, ByVal aCode As String)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aChoice = "Add" Then
            oTest.DoQuery("Delete  from [@Z_ADC2] where Name like '%_XD' and U_Z_AddRef='" & aCode & "'")
        Else
            oTest.DoQuery("Update [@Z_ADC2] set Name=Code where Name  Like '%_XD' and U_Z_AddRef='" & aCode & "'")
        End If
    End Sub
    Private Sub DeleteRow(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        oGrid = aForm.Items.Item("27").Specific
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) = True Then
                If oGrid.DataTable.GetValue("Code", intRow) <> "" Then
                    oTest.DoQuery("Update [@Z_ADC2] set Name = Name + '_XD' where Code='" & oGrid.DataTable.GetValue("Code", intRow) & "'")
                    oGrid.DataTable.Rows.Remove(intRow)
                    For intRow1 As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        oGrid.RowHeaders.SetText(intRow1, intRow1 + 1)
                    Next
                    aForm.Freeze(False)
                    Exit Sub
                End If
            End If
        Next
        aForm.Freeze(False)
    End Sub


    Private Sub formatgrid(ByVal aform As SAPbouiCOM.Form)
        oGrid = aform.Items.Item("27").Specific
        oGrid.Columns.Item("Code").Visible = False
        oGrid.Columns.Item("Name").Visible = False
        oGrid.Columns.Item("U_Z_AddRef").Visible = False
        oGrid.Columns.Item("U_Z_ItemCode").TitleObject.Caption = "Item Code"
        oEditTextColumn = oGrid.Columns.Item("U_Z_ItemCode")
        oEditTextColumn.ChooseFromListUID = "CFL_1"
        oEditTextColumn.ChooseFromListAlias = "ItemCode"
        oEditTextColumn.LinkedObjectType = "4"
        oGrid.Columns.Item("U_Z_ItemName").TitleObject.Caption = "Description"
        oGrid.Columns.Item("U_Z_ItemName").Editable = False
        oGrid.Columns.Item("U_Z_Price").TitleObject.Caption = "Unit Price"
        oGrid.Columns.Item("U_Z_Quantity").TitleObject.Caption = "Quantity"
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intRow, intRow + 1)

        Next
        oGrid.AutoResizeColumns()
    End Sub
    Private Sub Calcualtion(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        oCheckbox = aform.Items.Item("10").Specific
        oCheckbox1 = aform.Items.Item("12").Specific
        Dim dblPumboprice, dblWaterProofprice, dblDefTemp, dblTemp, dblTempPrice, dblTotalprice, dblUnitPrice As Double
        dblUnitPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "8"))
        If oCheckbox.Checked = True Then
            dblPumboprice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "18"))
        Else
            dblPumboprice = 0
        End If
        If oCheckbox1.Checked = True Then
            dblWaterProofprice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "19"))
        Else
            dblWaterProofprice = 0
        End If

        dblDefTemp = oApplication.Utilities.getEdittextvalue(aform, "22")
        If oApplication.Utilities.getEdittextvalue(aform, "14") = "" Then
            dblTemp = 0
        Else
            dblTemp = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "14"))
        End If
        dblWaterProofprice = 0
        dblTempPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "24"))
        dblTotalprice = (dblUnitPrice - dblPumboprice + dblWaterProofprice + (dblTempPrice * (dblDefTemp - dblTemp)))
        ' oApplication.Utilities.SetEditText(aform, "16", dblTotalprice)
        oGrid = aform.Items.Item("27").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue("U_Z_ItemCode", intRow) <> "" Then
                dblTotalprice = dblTotalprice + (oGrid.DataTable.GetValue("U_Z_Price", intRow) * oGrid.DataTable.GetValue("U_Z_Quantity", intRow))
            End If
        Next
        oApplication.Utilities.SetEditText(aform, "16", dblTotalprice)
        aform.Freeze(False)
    End Sub

#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim oRec As SAPbobsCOM.Recordset
        Dim strPumpLine As String = ""
        Dim strWaterproofLine As String = ""
        Dim strCode, strPump, strWaterproof, strEname, strETax, strGLAcc As String
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_ADC1")
        oCheckbox = aform.Items.Item("10").Specific
        oCheckbox1 = aform.Items.Item("12").Specific
        strCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        Dim dblPumboprice, dblWaterProofprice, dblDefTemp, dblTemp, dblTempPrice, dblTotalprice, dblUnitPrice As Double

        dblUnitPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "8"))
        If oCheckbox.Checked = True Then
            dblPumboprice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "18"))
        Else
            dblPumboprice = 0
        End If
        If oCheckbox1.Checked = True Then
            dblWaterProofprice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "19"))
        Else
            dblWaterProofprice = 0
        End If
        dblWaterProofprice = 0

        dblDefTemp = oApplication.Utilities.getEdittextvalue(aform, "22")
        If oApplication.Utilities.getEdittextvalue(aform, "14") = "" Then
            dblTemp = 0
        Else
            dblTemp = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "14"))
        End If

        dblTempPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "24"))
        dblTotalprice = (dblUnitPrice - dblPumboprice + dblWaterProofprice + (dblTempPrice * (dblDefTemp - dblTemp)))
        oApplication.Utilities.SetEditText(aform, "16", dblTotalprice)


        oGrid = aform.Items.Item("27").Specific
        oUserTable = oApplication.Company.UserTables.Item("Z_ADC2")
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue("U_Z_ItemCode", intRow) <> "" Then
                If oUserTable.GetByKey(oGrid.DataTable.GetValue("Code", intRow)) Then
                    oUserTable.Code = oGrid.DataTable.GetValue("Code", intRow)
                    oUserTable.Name = oGrid.DataTable.GetValue("Name", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = oGrid.DataTable.GetValue("U_Z_ItemCode", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = oGrid.DataTable.GetValue("U_Z_ItemName", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Price").Value = oGrid.DataTable.GetValue("U_Z_Price", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Quantity").Value = oGrid.DataTable.GetValue("U_Z_Quantity", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_AddRef").Value = oApplication.Utilities.getEdittextvalue(aform, "4")
                    dblTotalprice = dblTotalprice + (oGrid.DataTable.GetValue("U_Z_Price", intRow) * oGrid.DataTable.GetValue("U_Z_Quantity", intRow))
                    If oUserTable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Else
                    strCode = oApplication.Utilities.getMaxCode("@Z_ADC2", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = oGrid.DataTable.GetValue("U_Z_ItemCode", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = oGrid.DataTable.GetValue("U_Z_ItemName", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Price").Value = oGrid.DataTable.GetValue("U_Z_Price", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Quantity").Value = oGrid.DataTable.GetValue("U_Z_Quantity", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_AddRef").Value = oApplication.Utilities.getEdittextvalue(aform, "4")
                    dblTotalprice = dblTotalprice + (oGrid.DataTable.GetValue("U_Z_Price", intRow) * oGrid.DataTable.GetValue("U_Z_Quantity", intRow))
                    If oUserTable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
        Next
        oApplication.Utilities.SetEditText(aform, "16", dblTotalprice)

        oUserTable = oApplication.Company.UserTables.Item("Z_ADC1")
        If oUserTable.GetByKey(strCode) Then
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = oApplication.Utilities.getEdittextvalue(aform, "6")
            oUserTable.UserFields.Fields.Item("U_Z_UnitPrice").Value = oApplication.Utilities.getEdittextvalue(aform, "8")
            If oCheckbox.Checked = True Then
                strPump = "Y"
                strPumpLine = "Yes"
            Else
                strPump = "N"
                strPumpLine = "No"
            End If
            oUserTable.UserFields.Fields.Item("U_Z_WhoutPump").Value = strPump
            If oCheckbox1.Checked = True Then
                strWaterproof = "Y"
                strWaterproofLine = "Yes"
            Else
                strWaterproof = "N"
                strWaterproofLine = "No"
            End If
            oUserTable.UserFields.Fields.Item("U_Z_WaterProof").Value = strWaterproof
            oUserTable.UserFields.Fields.Item("U_Z_Temp").Value = oApplication.Utilities.getEdittextvalue(aform, "14")
            dblTotalprice = oApplication.Utilities.getEdittextvalue(aform, "16")
            oUserTable.UserFields.Fields.Item("U_Z_ActualPrice").Value = dblTotalprice ' oApplication.Utilities.getEdittextvalue(aform, "16")
            oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = oApplication.Utilities.getEdittextvalue(aform, "4")
            oUserTable.UserFields.Fields.Item("U_Z_WhPumpPrice").Value = oApplication.Utilities.getEdittextvalue(aform, "18")
            oUserTable.UserFields.Fields.Item("U_Z_WatProofPrice").Value = oApplication.Utilities.getEdittextvalue(aform, "19")
            oUserTable.UserFields.Fields.Item("U_Z_DefTemp").Value = oApplication.Utilities.getEdittextvalue(aform, "22")
            oUserTable.UserFields.Fields.Item("U_Z_TempValue").Value = oApplication.Utilities.getEdittextvalue(aform, "24")
            If oUserTable.Update() <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

        End If
        commintTrans("Add", oApplication.Utilities.getEdittextvalue(aform, "4"))
        Dim Rowid As Integer = oApplication.Utilities.getEdittextvalue(aform, "17")
        oMatrix = objSourceForm.Items.Item("38").Specific
        oApplication.Utilities.SetMatrixValues(oMatrix, "14", Rowid, oApplication.Utilities.getEdittextvalue(aform, "16"))
        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_WhoutPump", Rowid, strPumpLine)
        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_WaterProof", Rowid, strWaterproofLine)
        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Temp", Rowid, oApplication.Utilities.getEdittextvalue(aform, "14"))
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        oForm.Close()
    End Function
#End Region

    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim ActTemp, DefTemp, SetupTemp, Totaltemp, TempTotalPrice, UnitPrice As Integer
        ActTemp = oApplication.Utilities.getEdittextvalue(aform, "14")
        DefTemp = oApplication.Utilities.getEdittextvalue(aform, "22")
        If ActTemp > DefTemp Then
            oApplication.Utilities.Message("Temprature should  be Less than or Equal to the Default Temprature", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Calcualtion(aform)
        Return True
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_AddCostDetails Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "2" Then
                                    commintTrans("Cancel", oApplication.Utilities.getEdittextvalue(oForm, "4"))
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        If Validation(oForm) = True Then
                                            AddtoUDT1(oForm)
                                        Else
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Case "10"
                                        Calcualtion(oForm)
                                    Case "12"
                                        Calcualtion(oForm)
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "14" And pVal.CharPressed = 9 Then
                                    Calcualtion(oForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val, val1 As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "27" And pVal.ColUID = "U_Z_ItemCode" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            val1 = oDataTable.GetValue("ItemName", 0)
                                            oGrid = oForm.Items.Item("27").Specific
                                            Try
                                                oGrid.DataTable.SetValue("U_Z_Price", pVal.Row, oDataTable.GetValue("U_Z_ItemCost", 0))
                                                oGrid.DataTable.SetValue("U_Z_ItemName", pVal.Row, val1)
                                                oGrid.DataTable.SetValue("U_Z_Quantity", pVal.Row, oDataTable.GetValue("U_Z_DefQty", 0))
                                                oGrid.DataTable.SetValue("U_Z_ItemCode", pVal.Row, val)
                                            Catch ex As Exception

                                            End Try

                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try
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

                Case mnu_ADD_ROW
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        addrow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    If pVal.BeforeAction = True Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        DeleteRow(oForm)
                        BubbleEvent = False
                        Exit Sub
                    End If
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
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                'Select Case pVal.MenuUID
                '    Case mnu_LeaveMaster
                '        oMenuobject = New clsEarning
                '        oMenuobject.MenuEvent(pVal, BubbleEvent)
                'End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
