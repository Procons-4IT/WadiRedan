
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
    Public Sub LoadForm(ByVal objSourceForm As SAPbouiCOM.Form, ByVal Rowid As String, ByVal RefCode As String)
        oForm = oApplication.Utilities.LoadForm(xml_AddCostDetails, frm_AddCostDetails)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Dim oRec As SAPbobsCOM.Recordset
        Dim strqry As String
        oCheckbox = oForm.Items.Item("10").Specific
        oCheckbox1 = oForm.Items.Item("12").Specific
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
            oApplication.Utilities.setEdittextvalue(oForm, "16", oRec.Fields.Item("U_Z_ActualPrice").Value)
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
        End If
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

        dblTempPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "24"))
        dblTotalprice = (dblUnitPrice - dblPumboprice + dblWaterProofprice + (dblTempPrice * (dblDefTemp - dblTemp)))
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

        dblDefTemp = oApplication.Utilities.getEdittextvalue(aform, "22")
        If oApplication.Utilities.getEdittextvalue(aform, "14") = "" Then
            dblTemp = 0
        Else
            dblTemp = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "14"))
        End If

        dblTempPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "24"))
        dblTotalprice = (dblUnitPrice - dblPumboprice + dblWaterProofprice + (dblTempPrice * (dblDefTemp - dblTemp)))
        oApplication.Utilities.SetEditText(aform, "16", dblTotalprice)

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
            oUserTable.UserFields.Fields.Item("U_Z_ActualPrice").Value = oApplication.Utilities.getEdittextvalue(aform, "16")
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
                                Dim sCHFL_ID, val As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        oForm.Update()
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
                Case mnu_DELETE_ROW
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
