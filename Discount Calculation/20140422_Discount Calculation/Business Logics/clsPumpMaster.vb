Public Class clsPumpMaster
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCheckbox As SAPbouiCOM.CheckBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oCheckboxColumn As SAPbouiCOM.CheckBoxColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private MatrixId As String
    Private RowtoDelete As Integer
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_PumpMaster, frm_PumpMaster)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        Databind(oForm)
        oForm.Freeze(False)
    End Sub
    Private Sub Databind(ByVal sform As SAPbouiCOM.Form)
        Dim strQry As String
        Try
            sform.Freeze(True)
      
            oGrid = sform.Items.Item("3").Specific
            strQry = "Select Code,Name,U_Z_PumpCode,U_Z_PumpDesc,U_Z_InActive,U_Z_FromDate,U_Z_ToDate,U_Z_Active from [@Z_OPUMP] where Code=Name"
            oGrid.DataTable.ExecuteQuery(strQry)
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.Columns.Item("Code").TitleObject.Caption = "Code"
            oGrid.Columns.Item("Code").Visible = False
            oGrid.Columns.Item("Name").TitleObject.Caption = "Name"
            oGrid.Columns.Item("Name").Visible = False
            oGrid.Columns.Item("U_Z_PumpCode").TitleObject.Caption = "Pump Code"
            oGrid.Columns.Item("U_Z_PumpDesc").TitleObject.Caption = "Pump Name"
            oGrid.Columns.Item("U_Z_InActive").TitleObject.Caption = "In Active"
            oGrid.Columns.Item("U_Z_InActive").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("U_Z_FromDate").TitleObject.Caption = "From Date"
            oGrid.Columns.Item("U_Z_ToDate").TitleObject.Caption = "To Date"
            oGrid.Columns.Item("U_Z_Active").TitleObject.Caption = "Active"
            oGrid.Columns.Item("U_Z_Active").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("U_Z_Active")
            oComboColumn.ValidValues.Add("Y", "Yes")
            oComboColumn.ValidValues.Add("N", "No")
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("U_Z_Active").Visible = False
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            sform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            sform.Freeze(False)
        End Try
    End Sub
    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCo, strCo1, strCo2, strCo3, strCode1, strCode2, strCode3, strCode4, strFromDate, strToDate As String
        Dim FromDate, ToDate As Date
        oGrid = aform.Items.Item("3").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue("U_Z_PumpCode", intRow) <> "" Then
                If oGrid.DataTable.GetValue("U_Z_PumpDesc", intRow) = "" Then
                    oApplication.Utilities.Message("Enter Pump Description", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                oCheckboxColumn = oGrid.Columns.Item("U_Z_InActive")
                If oCheckboxColumn.IsChecked(intRow) = True Then
                    strFromDate = oGrid.DataTable.GetValue("U_Z_FromDate", intRow)
                    strToDate = oGrid.DataTable.GetValue("U_Z_ToDate", intRow)
                    If strFromDate = Nothing Then
                        oApplication.Utilities.Message("Enter InActive From Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strToDate = Nothing Then
                        oApplication.Utilities.Message("Enter InActive To Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    FromDate = oGrid.DataTable.GetValue("U_Z_FromDate", intRow)
                    ToDate = oGrid.DataTable.GetValue("U_Z_ToDate", intRow)
                    If FromDate > ToDate Then
                        oGrid.Columns.Item("U_Z_FromDate").Click(intRow, False, 0)
                        oApplication.Utilities.Message("To Date should be greater than From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If

              
                strCo3 = oGrid.DataTable.GetValue("U_Z_PumpCode", intRow)
                For intLoop As Integer = intRow + 1 To oGrid.DataTable.Rows.Count - 1
                    strCode4 = oGrid.DataTable.GetValue("U_Z_PumpCode", intLoop)
                    If strCode4 <> "" Then
                        If strCo3.ToUpper = strCode4.ToUpper Then
                            oApplication.Utilities.Message("Pump Code : This entry already exists : " & strCode4, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oGrid.Columns.Item("U_Z_PumpCode").Click(intLoop)
                            Return False
                        End If
                    End If
                Next
            End If
        Next
        Return True
    End Function
    
    Private Function AddUDT(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strTable, strCode, strType, strfromDate, strTodate As String
        Dim Fromdate, ToDate As Date
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oUsertable As SAPbobsCOM.UserTable
        oGrid = aform.Items.Item("3").Specific
        strTable = "@Z_OPUMP"
        oUsertable = oApplication.Company.UserTables.Item("Z_OPUMP")
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strCode = oGrid.DataTable.GetValue("Code", intRow)
            oEditTextColumn = oGrid.Columns.Item("U_Z_PumpCode")
            Try
                strType = oEditTextColumn.GetText(oGrid.DataTable.Rows.Count - 1).ToString
            Catch ex As Exception
                strType = ""
            End Try

            If strType <> "" Then
                If oUsertable.GetByKey(strCode) Then
                    oUsertable.Code = strCode
                    oUsertable.Name = strCode
                    oUsertable.UserFields.Fields.Item("U_Z_PumpCode").Value = oGrid.DataTable.GetValue("U_Z_PumpCode", intRow)
                    oUsertable.UserFields.Fields.Item("U_Z_PumpDesc").Value = oGrid.DataTable.GetValue("U_Z_PumpDesc", intRow)
                    strfromDate = oGrid.DataTable.GetValue("U_Z_FromDate", intRow)
                    If strfromDate <> Nothing Then
                        oUsertable.UserFields.Fields.Item("U_Z_FromDate").Value = oGrid.DataTable.GetValue("U_Z_FromDate", intRow)
                    Else
                        oUsertable.UserFields.Fields.Item("U_Z_FromDate").Value = ""
                    End If
                    strTodate = oGrid.DataTable.GetValue("U_Z_ToDate", intRow)
                    If strTodate <> Nothing Then
                        oUsertable.UserFields.Fields.Item("U_Z_ToDate").Value = oGrid.DataTable.GetValue("U_Z_ToDate", intRow)
                    Else
                        oUsertable.UserFields.Fields.Item("U_Z_ToDate").Value = ""
                    End If
                    oCheckboxColumn = oGrid.Columns.Item("U_Z_InActive")
                    If oCheckboxColumn.IsChecked(intRow) = True Then
                        oUsertable.UserFields.Fields.Item("U_Z_InActive").Value = "Y"
                    Else
                        oUsertable.UserFields.Fields.Item("U_Z_InActive").Value = "N"
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_Active").Value = "Y"
                    If oUsertable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Else
                    strCode = oApplication.Utilities.getMaxCode(strTable, "Code")
                    oUsertable.Code = strCode
                    oUsertable.Name = strCode
                    oUsertable.UserFields.Fields.Item("U_Z_PumpCode").Value = oGrid.DataTable.GetValue("U_Z_PumpCode", intRow)
                    oUsertable.UserFields.Fields.Item("U_Z_PumpDesc").Value = oGrid.DataTable.GetValue("U_Z_PumpDesc", intRow)
                    strfromDate = oGrid.DataTable.GetValue("U_Z_FromDate", intRow)
                    If strfromDate <> Nothing Then
                        oUsertable.UserFields.Fields.Item("U_Z_FromDate").Value = oGrid.DataTable.GetValue("U_Z_FromDate", intRow)
                    Else
                        oUsertable.UserFields.Fields.Item("U_Z_FromDate").Value = ""
                    End If
                    strTodate = oGrid.DataTable.GetValue("U_Z_ToDate", intRow)
                    If strTodate <> Nothing Then
                        oUsertable.UserFields.Fields.Item("U_Z_ToDate").Value = oGrid.DataTable.GetValue("U_Z_ToDate", intRow)
                    Else
                        oUsertable.UserFields.Fields.Item("U_Z_ToDate").Value = ""
                    End If
                    oCheckboxColumn = oGrid.Columns.Item("U_Z_InActive")
                    If oCheckboxColumn.IsChecked(intRow) = True Then
                        oUsertable.UserFields.Fields.Item("U_Z_InActive").Value = "Y"
                    Else
                        oUsertable.UserFields.Fields.Item("U_Z_InActive").Value = "N"
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_Active").Value = "Y"
                    If oUsertable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
        Next
        oRec.DoQuery("Delete from [@Z_OPUMP] where Name like '%_XD'")
        Databind(aform)
        Return True
    End Function
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("3").Specific
        oEditTextColumn = oGrid.Columns.Item("U_Z_PumpCode")

        Dim strCode As String
        If oGrid.DataTable.Rows.Count - 1 <= 0 Then
            oGrid.DataTable.Rows.Add()
        End If
        strCode = oEditTextColumn.GetText(oGrid.DataTable.Rows.Count - 1).ToString
        If strCode <> "" Then
            oGrid.DataTable.Rows.Add()
            If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End If
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next
        oGrid.Columns.Item("U_Z_PumpCode").Click(oGrid.DataTable.Rows.Count - 1)
    End Sub
#Region "DeleteRow"
    Private Sub DeleteRow(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("3").Specific
        Dim strCode, strDocEntry As String
        Dim oTemp, oRec As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                oRec.DoQuery("Select * from  ORDR where U_Z_PumpCode='" & strCode & "'")
                If oRec.RecordCount > 0 Then
                    oApplication.Utilities.Message("Pump Code already mapped in Sales Order", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                Else
                    oTemp.DoQuery("Update [@Z_OPUMP] set Name=Name+'_XD' where Code='" & strCode & "'")
                    oGrid.DataTable.Rows.Remove(intRow)
                    If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    For intRow1 As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        oGrid.RowHeaders.SetText(intRow1, intRow1 + 1)
                    Next
                    Exit Sub
                End If
            End If
        Next
    End Sub

    Private Sub CommitTrans(ByVal aform As SAPbouiCOM.Form)
        Dim otes As SAPbobsCOM.Recordset
        otes = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otes.DoQuery("Update  [@Z_OPUMP] set Name=Code where Name like '%_XD'")
    End Sub
#End Region
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PumpMaster Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim strCode As String
                                Dim oRec As SAPbobsCOM.Recordset
                                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If pVal.CharPressed <> 9 And pVal.ColUID = "U_Z_PumpCode" Then
                                    oGrid = oForm.Items.Item("3").Specific
                                    strCode = oGrid.DataTable.GetValue("Code", pVal.Row)
                                    If strCode <> "" Then
                                        oRec.DoQuery("Select * from  ORDR where U_Z_PumpCode='" & strCode & "'")
                                        If oRec.RecordCount > 0 Then
                                            oApplication.Utilities.Message("Pump Code already mapped in Sales Order", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If pVal.ItemUID = "2" Then
                                        CommitTrans(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And pVal.ColUID = "U_Z_InActive" Then
                                    oGrid = oForm.Items.Item("3").Specific
                                    Try
                                        oForm.Freeze(True)
                                        oCheckboxColumn = oGrid.Columns.Item("U_Z_InActive")
                                        If oCheckboxColumn.IsChecked(pVal.Row) = True Then
                                            oGrid.DataTable.SetValue("U_Z_FromDate", pVal.Row, "")
                                            oGrid.DataTable.SetValue("U_Z_ToDate", pVal.Row, "")
                                        End If
                                        oForm.Freeze(False)
                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                    End Try
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
                                    Case "11"
                                        If Validation(oForm) = True Then
                                            If AddUDT(oForm) = True Then
                                                oApplication.Utilities.Message("Operation Completed Successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If
                                        Else
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                    Case "4"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            AddRow(oForm)
                                        End If
                                    Case "5"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            DeleteRow(oForm)
                                        End If
                                End Select
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
                Case mnu_PumpMaster
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
