Public Class clsInvoice
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox, oCombobox1 As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oItem1 As SAPbouiCOM.Item
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Methods"
    Private Sub FillPump(ByVal sform As SAPbouiCOM.Form)
        oCombobox = sform.Items.Item("edComp").Specific
        Dim oSlpRS As SAPbobsCOM.Recordset
        oSlpRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            Try
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try
        Next
        oSlpRS.DoQuery("Select Code,U_Z_PumpDesc from [@Z_OPUMP] order by Code")
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 0 To oSlpRS.RecordCount - 1
            oCombobox.ValidValues.Add(oSlpRS.Fields.Item(0).Value, oSlpRS.Fields.Item(1).Value)
            oSlpRS.MoveNext()
        Next
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        sform.Items.Item("edComp").DisplayDesc = True
    End Sub

    Private Sub Addcontrols(ByVal aForm As SAPbouiCOM.Form)

        If aForm.TypeEx = frm_Order Then
            oApplication.Utilities.AddControls(aForm, "btnAd", "29", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "DOWN", 0, 0, , "Additional Cost", 120, 5, 5)
            oApplication.Utilities.AddControls(aForm, "btnDup", "btnAd", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "DOWN", 0, 0, , "Duplicate Sales Order", 130, 5, 5)
        End If

        If aForm.TypeEx = frm_SalesQuotation Then
            oApplication.Utilities.AddControls(aForm, "btnAd", "29", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "DOWN", 0, 0, , "Additional Cost", 120, 5, 5)
            '   oApplication.Utilities.AddControls(aForm, "btnDup", "btnAd", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "DOWN", 0, 0, , "Duplicate Sales Order", 130, 5, 5)
        End If

        Dim strTableName As String
        Select Case aForm.TypeEx
            Case frm_Order
                strTableName = "ORDR"
            Case frm_Delivery
                strTableName = "ODLN"
            Case frm_Invoice
                strTableName = "OINV"
            Case frm_ResInvoice
                strTableName = "OINV"
            Case frm_ARDownpayment
                strTableName = "ODPI"
            Case frm_SalesQuotation
                strTableName = "OQUT"

        End Select
      
        oApplication.Utilities.AddControls(aForm, "stcomp", "86", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "Pump Name", 120)
        oApplication.Utilities.AddControls(aForm, "edComp", "46", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "DOWN", 0, 0, , , 120)
        oCombobox = aForm.Items.Item("edComp").Specific
        oCombobox.DataBind.SetBound(True, strTableName, "U_Z_PumpCode")
        ' oEditText.DataBind.SetBound(True, "ORDR", "U_Z_PumpName")
        If strTableName <> "ORDR" Then
            aForm.Items.Item("edComp").Enabled = False
        Else
            aForm.Items.Item("edComp").Enabled = True
        End If
        oItem1 = aForm.Items.Item("stcomp")
        oItem1.LinkTo = "edComp"
        oApplication.Utilities.AddControls(aForm, "edpuname", "46", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", 0, 0, , , 20)
        oEditText = oForm.Items.Item("edpuname").Specific
        oEditText.DataBind.SetBound(True, strTableName, "U_Z_PumpName")
        aForm.Items.Item("edpuname").Visible = False

        oApplication.Utilities.AddControls(aForm, "stfrTime", "stcomp", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "From Time", 120)
        oApplication.Utilities.AddControls(aForm, "edfrTime", "edComp", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, , , 120)
        oEditText = oForm.Items.Item("edfrTime").Specific
        oEditText.DataBind.SetBound(True, strTableName, "U_Z_FromTime")
        oItem1 = aForm.Items.Item("stfrTime")
        oItem1.LinkTo = "edfrTime"

        oApplication.Utilities.AddControls(aForm, "edfrdate", "edComp", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", 0, 0, , , 20)
        oEditText = oForm.Items.Item("edfrdate").Specific
        oEditText.DataBind.SetBound(True, strTableName, "U_Z_FromDate")
        aForm.Items.Item("edfrdate").Visible = False

        oApplication.Utilities.AddControls(aForm, "stToTime", "stfrTime", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "To Time", 120)
        oApplication.Utilities.AddControls(aForm, "edToTime", "edfrTime", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, , , 120)
        oEditText = aForm.Items.Item("edToTime").Specific
        oEditText.DataBind.SetBound(True, strTableName, "U_Z_ToTime")
        oItem1 = aForm.Items.Item("stToTime")
        oItem1.LinkTo = "edToTime"

        oApplication.Utilities.AddControls(aForm, "edtodate", "edfrTime", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", 0, 0, , , 20)
        oEditText = oForm.Items.Item("edtodate").Specific
        oEditText.DataBind.SetBound(True, strTableName, "U_Z_ToDate")
        aForm.Items.Item("edtodate").Visible = False
    End Sub

#End Region


    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim fromTime, ToTime, strstatus, strPunpcode As String
        Dim deldate As Date
        Dim strFrTime, strToTime, strdeldate As String
        Dim oRec, oRec1, oTemp As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            aform.Freeze(True)
            If aform.TypeEx = frm_Order Or aform.TypeEx = frm_Delivery Or aform.TypeEx = frm_Invoice Then
                oMatrix = aform.Items.Item("38").Specific
                For intRow As Integer = 1 To oMatrix.RowCount
                    oApplication.Utilities.CalculatePallet(aform, intRow)
                Next
            End If
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try

        If aform.TypeEx <> frm_Order Then
            Return True
        End If



        fromTime = oApplication.Utilities.getEdittextvalue(aform, "edfrTime")
        ToTime = oApplication.Utilities.getEdittextvalue(aform, "edToTime")
        strdeldate = oApplication.Utilities.getEdittextvalue(aform, "12")

        Try
            deldate = oApplication.Utilities.GetDateTimeValue(strdeldate)
        Catch ex As Exception
            deldate = Now.Date
        End Try

        oCombobox = aform.Items.Item("81").Specific
        strstatus = oCombobox.Selected.Value
        oCombobox1 = aform.Items.Item("edComp").Specific
        strPunpcode = oCombobox1.Selected.Value
        If strstatus = "3" Then
            ' oApplication.Utilities.Message("Document is Closed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If strPunpcode <> "" Then
            oRec.DoQuery("Select * from [@Z_OPUMP] where  Code='" & strPunpcode & "' and U_Z_InActive='Y'")
            If oRec.RecordCount > 0 Then
                oTemp.DoQuery("Select * from [@Z_OPUMP] where '" & deldate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate and Code='" & strPunpcode & "'")
                If oTemp.RecordCount > 0 Then
                    oApplication.Utilities.Message("Selected Pump is InActive...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            If fromTime = "" Then
                oApplication.Utilities.Message("Fromtime is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf ToTime = "" Then
                oApplication.Utilities.Message("Totime  is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                strFrTime = deldate.ToString("yyyy-MM-dd") & " " & fromTime
                strToTime = deldate.ToString("yyyy-MM-dd") & " " & ToTime
                oApplication.Utilities.setEdittextvalue(aform, "edfrdate", strFrTime)
                oApplication.Utilities.setEdittextvalue(aform, "edtodate", strToTime)
                If validateGateOutTime(deldate.ToString("yyyy-MM-dd"), fromTime, deldate.ToString("yyyy-MM-dd"), ToTime) <= 0 Then
                    oApplication.Utilities.Message("To time should be greater than From time...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                strSQL = "Select * from ORDR where  docnum <>" & oApplication.Utilities.getEdittextvalue(aform, "8") & " and  DocStatus='O' and U_Z_PumpCode='" & strPunpcode & "' and ( '" & strFrTime & "' between U_Z_FromDate and U_Z_ToDate)"
                'strSQL = "Select * from ORDR where DocStatus='O' and U_Z_PumpCode='" & strPunpcode & "' and  U_Z_Fromdate = '" & oApplication.Utilities.getEdittextvalue(aform, "edfrdate") & "' and U_Z_Todate='" & oApplication.Utilities.getEdittextvalue(aform, "edtodate") & "'"
                oRec1.DoQuery(strSQL)
                If oRec1.RecordCount > 0 Then
                    oApplication.Utilities.Message("Selected Pump already allocated to another document....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
        End If
        Return True
    End Function
    Public Function validateGateOutTime(ByVal aDate As String, ByVal aTime As String, ByVal aToDate As String, ByVal aToTime As String) As Double
        Dim strFromdate, strTodate, strFromTime, strtoTime, strDt1, strDt2 As String
        Dim dtFrom, dtto As Date
        Dim dtFromtime, dtToTime As Integer
        Try
            strDt1 = aDate ' oApplication.Utilities.getEdittextvalue(aForm, "7")
            strDt2 = aToDate ' oApplication.Utilities.getEdittextvalue(aForm, "159")
            strFromTime = aTime ' oApplication.Utilities.getEdittextvalue(aForm, "12")
            strtoTime = aToTime 'oApplication.Utilities.getEdittextvalue(aForm, "160")
            dtFrom = strDt1 ' oApplication.Utilities.getEdittextvalue(aForm, "7")
            dtto = strDt2 ' oApplication.Utilities.getEdittextvalue(aForm, "159")
            If strDt1 <> "" And strDt2 <> "" Then
                If strFromTime <> "" Then
                    strFromdate = dtFrom.ToString("yyyy-MM-dd") & " " & strFromTime & ":00"
                Else
                    strFromdate = dtFrom.ToString("yyyy-MM-dd")
                End If

                If strtoTime <> "" Then
                    strtoTime = dtto.ToString("yyyy-MM-dd") & " " & strtoTime & ":00"
                Else
                    strtoTime = dtto.ToString("yyyy-MM-dd")
                End If
            End If

            Dim oTest As SAPbobsCOM.Recordset
            Dim strsql As String
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strsql = "Select DateDiff(hour,'" & strFromdate & "','" & strtoTime & "')"
            oTest.DoQuery(strsql)
            Dim dblDifference As Double
            Dim intDiffer As Integer
            dblDifference = oTest.Fields.Item(0).Value
            dblDifference = 0
            If (dblDifference = 0) Then
                strsql = "Select DateDiff(minute,'" & strFromdate & "','" & strtoTime & "')"
                oTest.DoQuery(strsql)

                dblDifference = oTest.Fields.Item(0).Value
            End If
            dblDifference = dblDifference / 60
            Return dblDifference

        Catch ex As Exception

        End Try
    End Function
#Region "Item Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If oForm.TypeEx = frm_Order Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "AP"
                        oCreationPackage.String = "Pump Availabiltity Check"
                        oCreationPackage.Enabled = True
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)


                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oApplication.SBO_Application.Menus.RemoveEx("AP")
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalesQuotation Or pVal.FormTypeEx = frm_Order Or pVal.FormTypeEx = frm_Delivery Or pVal.FormTypeEx = frm_Invoice Or pVal.FormTypeEx = frm_ARDownpayment Or pVal.FormTypeEx = frm_ResInvoice Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "1" And pVal.FormTypeEx = frm_Delivery And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oApplication.Utilities.UpdateBOM(oMatrix) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" Then
                                    intselectedrow = pVal.Row
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_AddRef" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If pVal.ItemUID = "edfrTime" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edToTime" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edComp" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_AddRef" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_WhoutPump" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_WaterProof" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_Temp" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edfrTime" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edToTime" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edComp" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_AddRef" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If pVal.ItemUID = "edfrTime" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edToTime" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edComp" And pVal.FormTypeEx <> frm_Order Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Addcontrols(oForm)
                                FillPump(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'If pVal.ItemUID = "38" And pVal.ColUID = "11" And pVal.CharPressed = 9 Then
                                '    oApplication.Utilities.CalculatePallet(oForm, pVal.Row)
                                'End If
                                'If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_PALLET" Or pVal.ColUID = "U_Z_LAYER" Or pVal.ColUID = "U_Z_ECH") And pVal.CharPressed = 9 Then
                                '    oApplication.Utilities.CalculateQTY(oForm, pVal.Row)
                                'End If
                                'If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_DISCOUNT" Or pVal.ColUID = "U_TotalDis") And pVal.CharPressed = 9 Then
                                '    oApplication.Utilities.CalculateDiscount(oForm, pVal.Row, pVal.ColUID)
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'If pVal.ItemUID = "38" And pVal.ColUID = "11" And pVal.CharPressed = 9 Then
                                '    oApplication.Utilities.CalculatePallet(oForm, pVal.Row)
                                'End If
                                'If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_PALLET" Or pVal.ColUID = "U_Z_LAYER" Or pVal.ColUID = "U_Z_ECH") And pVal.CharPressed = 9 Then
                                '    oApplication.Utilities.CalculateQTY(oForm, pVal.Row)
                                'End If
                                'If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_DISCOUNT" Or pVal.ColUID = "U_TotalDis") And pVal.CharPressed = 9 Then
                                '    oApplication.Utilities.CalculateDiscount(oForm, pVal.Row, pVal.ColUID)
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "edComp" Then
                                    oCombobox = oForm.Items.Item("edComp").Specific
                                    oApplication.Utilities.setEdittextvalue(oForm, "edpuname", oCombobox.Selected.Description)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "btnDup" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    Dim strItemCode, Rowid, Qty, RefCode As String
                                    Dim Deldate As String
                                    RefCode = oApplication.Utilities.getEdittextvalue(oForm, "8")
                                    Dim oObj As New clsDupSalesOrder

                                    oObj.LoadForm(RefCode)
                                End If
                                If pVal.ItemUID = "btnAd" Then
                                    Dim strItemCode, Rowid, Qty, RefCode As String
                                    Dim Deldate As String
                                    Dim strUnitprice As Double
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oMatrix.IsRowSelected(intselectedrow) = False Then
                                        oApplication.Utilities.Message("No rows has been selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Else
                                        Deldate = oApplication.Utilities.getEdittextvalue(oForm, "12")
                                        If Deldate <> "" Then
                                            If 1 = 1 Then
                                                Rowid = intselectedrow
                                                strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", Rowid)
                                                strUnitprice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", Rowid))
                                                RefCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_AddRef", Rowid)
                                                If AddtoUDT(oForm, strItemCode, RefCode, Rowid, strUnitprice) = True Then
                                                    Dim AddCostDetails As New clsQCMaster
                                                    objSourceForm = oForm
                                                    RefCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_AddRef", Rowid)
                                                    AddCostDetails.LoadForm(objSourceForm, Rowid, RefCode)
                                                End If
                                            End If
                                        Else
                                            oApplication.Utilities.Message("Delivery date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If


                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim strItemCode, Rowid, Qty, RefCode As String
                                Dim Deldate As String
                                Dim strUnitprice As Double
                                oMatrix = oForm.Items.Item("38").Specific
                                Deldate = oApplication.Utilities.getEdittextvalue(oForm, "12")
                                If oForm.TypeEx <> frm_Order Then
                                    Exit Sub
                                End If
                                If Deldate <> "" Then
                                    If pVal.ItemUID = "38" And pVal.Row <> 0 And pVal.ColUID <> "0" Then
                                        Rowid = pVal.Row
                                        strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                        strUnitprice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", pVal.Row))
                                        RefCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_AddRef", pVal.Row)
                                        If AddtoUDT(oForm, strItemCode, RefCode, Rowid, strUnitprice) = True Then
                                            Dim AddCostDetails As New clsQCMaster
                                            objSourceForm = oForm
                                            RefCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_AddRef", pVal.Row)
                                            AddCostDetails.LoadForm(objSourceForm, Rowid, RefCode)
                                        End If
                                    End If
                                Else
                                    oApplication.Utilities.Message("Delivery date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1, Val7 As String
                                Dim sCHFL_ID, val, val2, val3, val4, val5, val6 As String
                                Dim intChoice As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "edComp" Then
                                            Try
                                                val = oDataTable.GetValue("U_Z_PumpCode", 0)
                                                val1 = oDataTable.GetValue("U_Z_PumpDesc", 0)
                                                oApplication.Utilities.setEdittextvalue(oForm, "edComp", val1)
                                                oApplication.Utilities.setEdittextvalue(oForm, "edpuname", val)
                                            Catch ex As Exception
                                                oForm.Freeze(False)
                                            End Try
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region
    Private Function AddtoUDT(ByVal aform As SAPbouiCOM.Form, ByVal ItemCode As String, ByVal RefCode As String, ByVal Rowid As Integer, ByVal strUnitprice As Double) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_ADC1")
        oMatrix = aform.Items.Item("38").Specific
        If RefCode <> "" Then
            strCode = RefCode
            'oUserTable.GetByKey(strCode)
            'oUserTable.Code = strCode
            'oUserTable.Name = strCode
            'oUserTable.Update()
        Else
            strCode = oApplication.Utilities.getMaxCode("@Z_ADC1", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
            oUserTable.UserFields.Fields.Item("U_Z_UnitPrice").Value = strUnitprice
            oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = strCode
            Dim dtdate As Date
            Try
                dtdate = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "12"))
            Catch ex As Exception
                dtdate = Now.Date
            End Try

            strqry = "Select * from [@Z_ADCS] where U_Z_Active='Y' and '" & dtdate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
            otemp.DoQuery(strqry)
            If otemp.RecordCount > 0 Then
                oUserTable.UserFields.Fields.Item("U_Z_WhoutPump").Value = "N"
                oUserTable.UserFields.Fields.Item("U_Z_WaterProof").Value = "N"
                oUserTable.UserFields.Fields.Item("U_Z_WhPumpPrice").Value = otemp.Fields.Item("U_Z_WhoutPump").Value
                oUserTable.UserFields.Fields.Item("U_Z_WatProofPrice").Value = otemp.Fields.Item("U_Z_WaterProof").Value
                oUserTable.UserFields.Fields.Item("U_Z_DefTemp").Value = otemp.Fields.Item("U_Z_DefTemp").Value
                oUserTable.UserFields.Fields.Item("U_Z_Temp").Value = otemp.Fields.Item("U_Z_DefTemp").Value
                oUserTable.UserFields.Fields.Item("U_Z_TempValue").Value = otemp.Fields.Item("U_Z_Temp").Value
                oUserTable.UserFields.Fields.Item("U_Z_ActualPrice").Value = 0
            Else
                oUserTable.UserFields.Fields.Item("U_Z_WhoutPump").Value = "N"
                oUserTable.UserFields.Fields.Item("U_Z_WaterProof").Value = "N"
                oUserTable.UserFields.Fields.Item("U_Z_WhPumpPrice").Value = 0 'otemp.Fields.Item("U_Z_WhoutPump").Value
                oUserTable.UserFields.Fields.Item("U_Z_WatProofPrice").Value = 0 'otemp.Fields.Item("U_Z_WaterProof").Value
                oUserTable.UserFields.Fields.Item("U_Z_DefTemp").Value = 0 'otemp.Fields.Item("U_Z_DefTemp").Value
                oUserTable.UserFields.Fields.Item("U_Z_Temp").Value = 0 ' otemp.Fields.Item("U_Z_DefTemp").Value
                oUserTable.UserFields.Fields.Item("U_Z_TempValue").Value = 0 ' otemp.Fields.Item("U_Z_Temp").Value
                oUserTable.UserFields.Fields.Item("U_Z_ActualPrice").Value = 0
            End If
            oUserTable.Add()
            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_AddRef", Rowid, strCode)
        End If

        Return True
    End Function

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "AP"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Dim ItemCode, ItemName As String

                    If pVal.BeforeAction = False Then
                        If oForm.TypeEx = frm_Order Then
                            If oApplication.Utilities.getEdittextvalue(oForm, "12") <> "" Then
                                ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "12")
                                ItemName = oApplication.Utilities.getEdittextvalue(oForm, "edfrTime")
                                If ItemName <> "" Then
                                    ItemName = oApplication.Utilities.getEdittextvalue(oForm, "edfrTime")
                                Else
                                    ItemName = "00:00"
                                End If
                                Dim dtDate As Date
                                dtDate = oApplication.Utilities.GetDateTimeValue(ItemCode)

                                ItemName = dtDate.ToString("yyyy-MM-dd") & " " & ItemName
                                Dim objct As New clsAvailabiltyCheck
                                objct.LoadForm(dtDate.ToString("yyyy-MM-dd"), ItemName)
                            Else
                                oApplication.Utilities.Message("Delivery Date is missing... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            oApplication.Utilities.Message("Delivery Date is missing... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_Order Then
                    oForm.Items.Item("edComp").Enabled = True
                ElseIf oForm.TypeEx = frm_Delivery Or oForm.TypeEx = frm_Invoice Then
                    oForm.Items.Item("edComp").Enabled = False
                End If
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Dim oobj As SAPbobsCOM.Documents
                Dim strcode As String = ""
                Dim strpumpcode, strpumpName, strstring, strDocCode As String
                Dim strfromhr, strtohr As String
                Dim otest1 As SAPbobsCOM.Recordset
                otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oApplication.Company.GetNewObjectCode(strcode)
                oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                If BusinessObjectInfo.FormTypeEx = frm_Delivery Then
                    oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        oApplication.Utilities.RemoveBoMItem(oobj.DocEntry)
                    End If
                End If
                If oForm.TypeEx = frm_Order Then
                    'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                    '        strDocCode = oApplication.Utilities.getMaxCode("@Z_OACAD", "Code")
                    '        oCombobox = oForm.Items.Item("edComp").Specific
                    '        strpumpcode = oCombobox.Selected.Value
                    '        strpumpName = oCombobox.Selected.Description
                    '        strfromhr = oApplication.Utilities.getEdittextvalue(oForm, "edfrTime")
                    '        strtohr = oApplication.Utilities.getEdittextvalue(oForm, "edToTime")
                    '        otest1.DoQuery("Update ORDR set U_Z_PumpName='" & strpumpName & "' where DocEntry=" & oobj.DocEntry)
                    '        strstring = "Insert into [@Z_OACAD] (Code,Name,U_Z_DocNum,U_Z_DocEntry,U_Z_DocDate,U_Z_DelDate,U_Z_PumpCode,U_Z_PumpName,U_Z_FromHour,U_Z_ToHour) "
                    '        strstring += " values ('" & strDocCode & "','" & strDocCode & "'," & oobj.DocEntry & "," & oobj.DocEntry & ",'" & oobj.DocDate.ToString("yyyy-MM-dd") & "','" & oobj.DocDueDate.ToString("yyyy-MM-dd") & "','" & strpumpcode & "','" & strpumpName & "','" & strfromhr.Replace(":", "") & "','" & strtohr.Replace(":", "") & "')"
                    '        Try
                    '            otest1.DoQuery(strstring)
                    '        Catch ex As Exception

                    '        End Try
                    '    End If
                Else
                    'If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                    '    oCombobox = oForm.Items.Item("edComp").Specific
                    '    strpumpcode = oCombobox.Selected.Value
                    '    strpumpName = oCombobox.Selected.Description
                    '    strfromhr = oApplication.Utilities.getEdittextvalue(oForm, "edfrTime")
                    '    strtohr = oApplication.Utilities.getEdittextvalue(oForm, "edToTime")
                    '    otest1.DoQuery("Update ORDR set U_Z_PumpName='" & strpumpName & "' where DocEntry=" & oobj.DocEntry)
                    '    strstring = "Update [@Z_OACAD] set U_Z_DocDate='" & oobj.DocDate.ToString("yyyy-MM-dd") & "',U_Z_DelDate='" & oobj.DocDueDate.ToString("yyyy-MM-dd") & "',U_Z_PumpCode='" & strpumpcode & "',U_Z_PumpName='" & strpumpName & "',U_Z_FromHour='" & strfromhr.Replace(":", "") & "',U_Z_ToHour='" & strtohr.Replace(":", "") & "' where U_Z_DocNum=" & oobj.DocEntry & " "

                    '    Try
                    '        otest1.DoQuery(strstring)
                    '    Catch ex As Exception

                    '    End Try
                    'End If
                    ' End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
