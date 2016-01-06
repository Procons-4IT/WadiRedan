Public Class clsAvailabiltyCheck
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
    Public Sub LoadForm(ByVal aDate As String, ByVal aTime As String)
        oForm = oApplication.Utilities.LoadForm(xml_AvCheck, frm_AvCheck)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
           
            Dim strquery, strquery1 As String
            'strquery = "select  X.U_Z_PumpCode,x.U_Z_PumpName, x.DocNum,X.CardCode,x.CardName,x.DocDueDate,x.U_Z_FromTime,x.U_Z_ToTime from "
            'strquery = strquery & "(Select DocDueDate  , U_Z_FromTime , U_Z_ToTime,DocNum,DocEntry,U_Z_PumpCode,U_Z_PumpName ,CardCode,CardName  from ORDR where DocStatus<>'C' and   U_Z_PumpCode<>'' and  ('" & aTime & "' between U_Z_FromDate and U_Z_ToDate)"
            'strquery = strquery & "    union"
            'strquery = strquery & " select  DocDueDate  ,U_Z_FromTime , U_Z_ToTime,DocNum,DocEntry,U_Z_PumpCode,U_Z_PumpName ,CardCode ,CardName  from ORDR where DocStatus<>'C' and  U_Z_PumpCode<>'' and convert(varchar(10),DocDueDate,104) > '" & aDate & "') x"

            strquery = "select  X.U_Z_PumpCode,x.U_Z_PumpName,X.DocEntry, x.DocNum,X.CardCode,x.CardName,x.DocDueDate,x.U_Z_FromTime,x.U_Z_ToTime,X.Dscription,X.U_Z_LOCATION,X.U_Z_Type,X.WhsCode,x.Quantity from "
            strquery = strquery & "(Select DocDueDate  , SUBSTRING(U_Z_FromDate,12,5) 'U_Z_FromTime',substring(U_Z_ToDate,12,5) 'U_Z_ToTime',DocNum,T0.DocEntry,U_Z_PumpCode,U_Z_PumpName ,T0.CardCode,CardName,Sum(T1.Quantity) 'Quantity',T1.WhsCode 'WhsCode',T1.Dscription,T1.U_Z_LOCATION ,T1.U_Z_TYPE  from ORDR T0 inner join RDR1 T1 on T0.DocEntry=T1.DocEntry inner Join OITM T2 on T2.ItemCode=T1.ItemCode where DocStatus<>'C' and   U_Z_PumpCode<>'' and  ('" & aTime & "' between U_Z_FromDate and U_Z_ToDate) group by T0.DocEntry,DocDueDate  , U_Z_FromDate , U_Z_ToDate,DocNum,U_Z_PumpCode,U_Z_PumpName ,T0.CardCode,CardName,T1.Dscription,T1.U_Z_LOCATION ,T1.U_Z_TYPE,WhsCode"
            strquery = strquery & "    union"
            ' strquery = strquery & " select  DocDueDate  ,SUBSTRING(U_Z_FromDate,12,5) 'U_Z_FromTime',substring(U_Z_ToDate,12,5) 'U_Z_ToTime',DocNum,T0.DocEntry,U_Z_PumpCode,U_Z_PumpName ,T0.CardCode ,CardName,Sum(T1.Quantity) 'Quantity',T1.WhsCode 'WhsCode' ,T1.Dscription,T1.U_Z_LOCATION ,T1.U_Z_TYPE from ORDR T0 inner join RDR1 T1 on T0.DocEntry=T1.DocEntry inner Join OITM T2 on T2.ItemCode=T1.ItemCode where DocStatus<>'C' and  U_Z_PumpCode<>'' and convert(varchar(10),DocDueDate,104) = '" & aDate & "' group by T0.DocEntry,DocDueDate  , U_Z_FromDate , U_Z_ToDate,DocNum,U_Z_PumpCode,U_Z_PumpName ,T0.CardCode,CardName,T1.Dscription,T1.U_Z_LOCATION ,T1.U_Z_TYPE,WhsCode) x order by X.DocNum"
            strquery = strquery & " select  DocDueDate  ,SUBSTRING(U_Z_FromDate,12,5) 'U_Z_FromTime',substring(U_Z_ToDate,12,5) 'U_Z_ToTime',DocNum,T0.DocEntry,U_Z_PumpCode,U_Z_PumpName ,T0.CardCode ,CardName,Sum(T1.Quantity) 'Quantity',T1.WhsCode 'WhsCode' ,T1.Dscription,T1.U_Z_LOCATION ,T1.U_Z_TYPE from ORDR T0 inner join RDR1 T1 on T0.DocEntry=T1.DocEntry inner Join OITM T2 on T2.ItemCode=T1.ItemCode where DocStatus<>'C' and  U_Z_PumpCode<>'' and DocDueDate = '" & aDate & "' group by T0.DocEntry,DocDueDate  , U_Z_FromDate , U_Z_ToDate,DocNum,U_Z_PumpCode,U_Z_PumpName ,T0.CardCode,CardName,T1.Dscription,T1.U_Z_LOCATION ,T1.U_Z_TYPE,WhsCode) x order by X.DocNum"
            oGrid = oForm.Items.Item("6").Specific
            oGrid.DataTable.ExecuteQuery(strquery)
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.Columns.Item(0).TitleObject.Caption = "Pump Code"
            oGrid.Columns.Item(0).Visible = False
            oGrid.Columns.Item(1).TitleObject.Caption = "Pump Name"
            oGrid.Columns.Item(2).TitleObject.Caption = "Document Entry"
            oEditTextColumn = oGrid.Columns.Item(2)
            oEditTextColumn.LinkedObjectType = "17"
            oGrid.Columns.Item(3).TitleObject.Caption = "Document Number"
            oGrid.Columns.Item(4).TitleObject.Caption = "Customer Code"
            oEditTextColumn = oGrid.Columns.Item(4)
            oEditTextColumn.LinkedObjectType = "2"
            oGrid.Columns.Item(5).TitleObject.Caption = "Customer Name"
            oGrid.Columns.Item(6).TitleObject.Caption = "Delivery Date"
            oGrid.Columns.Item(7).TitleObject.Caption = "From Time"
            oGrid.Columns.Item(8).TitleObject.Caption = "To Time"
            oGrid.Columns.Item("Dscription").TitleObject.Caption = "Item Name"
            'oEditTextColumn = oGrid.Columns.Item("Dscription")
            'oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("U_Z_LOCATION").TitleObject.Caption = "Location"
            oGrid.Columns.Item("U_Z_Type").TitleObject.Caption = "Type"
            oGrid.Columns.Item("WhsCode").TitleObject.Caption = "Warehouse"
            oEditTextColumn = oGrid.Columns.Item("WhsCode")
            oEditTextColumn.LinkedObjectType = "64"
            oGrid.Columns.Item("Quantity").TitleObject.Caption = "Total Quantity"
            oEditTextColumn = oGrid.Columns.Item(8)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.AutoResizeColumns()
            '  oGrid.CollapseLevel = 2
            strquery = " select * from [@Z_OPUMP] where Code not in (select  X.U_Z_PumpCode from "
            strquery = strquery & "(Select DocDueDate  , U_Z_FromDate , U_Z_ToDate,DocNum,DocEntry,U_Z_PumpCode,U_Z_PumpName   from ORDR where DocStatus<>'C' and   U_Z_PumpCode<>'' and  ('" & aTime & "' between U_Z_FromDate and U_Z_ToDate)"
            strquery = strquery & "    union"
            ' strquery = strquery & " select  DocDueDate  ,U_Z_FromDate , U_Z_ToDate,DocNum,DocEntry,U_Z_PumpCode,U_Z_PumpName    from ORDR where DocStatus<>'C' and  U_Z_PumpCode<>'' and convert(varchar(10),DocDueDate,104) > '" & aDate & "') x)"
            strquery = strquery & " select  DocDueDate  ,U_Z_FromDate , U_Z_ToDate,DocNum,DocEntry,U_Z_PumpCode,U_Z_PumpName    from ORDR where DocStatus<>'C' and  U_Z_PumpCode<>'' and DocDueDate > '" & aDate & "') x)"

            oGrid = oForm.Items.Item("5").Specific
            oGrid.DataTable.ExecuteQuery(strquery)
            oGrid.Columns.Item("Code").Visible = False
            oGrid.Columns.Item("Name").Visible = False
            oGrid.Columns.Item("U_Z_PumpCode").TitleObject.Caption = "Pump Code"
            oGrid.Columns.Item("U_Z_PumpDesc").TitleObject.Caption = "Pump Description"
            oGrid.Columns.Item("U_Z_InActive").TitleObject.Caption = "InActive Status"
            oGrid.Columns.Item("U_Z_InActive").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("U_Z_FromDate").TitleObject.Caption = "From Date"
            oGrid.Columns.Item("U_Z_ToDate").TitleObject.Caption = "To Date"
            oGrid.Columns.Item("U_Z_Active").Visible = False
            oGrid.AutoResizeColumns()
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            oForm.Freeze(False)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)

        End Try
    End Sub

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_AvCheck Then
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
                                        oForm.PaneLevel = 1
                                    Case "4"
                                        oForm.PaneLevel = 2

                                End Select


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

End Class
