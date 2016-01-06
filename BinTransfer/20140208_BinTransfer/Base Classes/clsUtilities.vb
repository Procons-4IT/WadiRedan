Public Class clsUtilities

    Private oMatrix As SAPbouiCOM.Matrix
    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub
#Region "Get Edit Text"
    Public Function getEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String) As String
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        Return objEdit.String
    End Function
    Public Sub setEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String, ByVal newvalue As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        objEdit.String = newvalue
    End Sub
#End Region
    Public Sub SetEditText(ByVal aForm As SAPbouiCOM.Form, ByVal aUID As String, ByVal aVal As String)
        Dim objedit As SAPbouiCOM.EditText
        objedit = aForm.Items.Item(aUID).Specific
        objedit.Value = aVal
    End Sub
    Public Sub createInventoryTransfer(ByVal aDocEntry As Integer, ByVal DocDate As Date, ByVal ItemCode As String, ByVal WhsCode As String, ByVal FrmBin As String, ByVal Quantity As Double, ByVal DocNum As Double, ByVal aLineNo As Integer)
        Dim transfer As SAPbobsCOM.StockTransfer
        Dim oRec As SAPbobsCOM.Recordset
        Dim blnLineexists As Boolean = False
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        transfer = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
        transfer.DocDate = DocDate
        transfer.FromWarehouse = WhsCode
        transfer.ToWarehouse = WhsCode
        transfer.Comments = "Created Based on Delivery : " & DocNum
        Dim s As String
        s = "SELECT T2.[U_Z_BinTransfer] ,T1.[Quantity],T1.[Code] FROM OITT T0  INNER JOIN ITT1 T1 ON T0.Code = T1.Father INNER JOIN OITM T2 ON T1.Code = T2.ItemCode WHERE T0.[Code] ='" & ItemCode & "' and T2.[U_Z_BinTransfer]='Y'"
        oRec.DoQuery(s)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            If intRow > 0 Then
                transfer.Lines.Add()
            End If
            Dim dblBomQty, dblDelQty, dblAllocateQTy As Double
            dblBomQty = oRec.Fields.Item("Quantity").Value
            dblDelQty = Quantity
            dblAllocateQTy = dblBomQty * dblDelQty

            transfer.Lines.SetCurrentLine(intRow)
            transfer.Lines.ItemCode = oRec.Fields.Item("Code").Value
            transfer.Lines.FromWarehouseCode = WhsCode
            transfer.Lines.WarehouseCode = WhsCode
            transfer.Lines.Quantity = dblAllocateQTy

            transfer.Lines.BinAllocations.SetCurrentLine(0)
            transfer.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batFromWarehouse
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("Select * from OBIN where BinCode='" & FrmBin & "' and whscode='" & WhsCode & "'")
            Dim inttobin As Integer = oTest.Fields.Item("AbsEntry").Value
            transfer.Lines.BinAllocations.BinAbsEntry = oTest.Fields.Item("AbsEntry").Value
            transfer.Lines.BinAllocations.Quantity = dblAllocateQTy
            transfer.Lines.BinAllocations.Add()

            transfer.Lines.BinAllocations.SetCurrentLine(1)
            transfer.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batToWarehouse
            oTest.DoQuery("Select * from OBIN where SysBin='Y' and WhsCode='" & WhsCode & "'")
            inttobin = oTest.Fields.Item("AbsEntry").Value
            transfer.Lines.BinAllocations.BinAbsEntry = oTest.Fields.Item("AbsEntry").Value
            transfer.Lines.BinAllocations.Quantity = dblAllocateQTy
            transfer.Lines.BinAllocations.Add()
            blnLineexists = True
            oRec.MoveNext()
        Next
        If blnLineexists = True Then
            If transfer.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Dim oTes As SAPbobsCOM.Recordset
                Dim strDoc As String
                oApplication.Company.GetNewObjectCode(strDoc)

                oTes = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTes.DoQuery("Select * from OWTR where DocEntry=" & CInt(strDoc))
                oTes.DoQuery("Update DLN1 set U_Z_ITNO='" & oTes.Fields.Item("DocNum").Value & "' where DocEntry=" & aDocEntry & " and LineNum=" & aLineNo)


                oApplication.Utilities.Message("Inventorty transfer added:", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
        End If
    End Sub

#Region "AddControls"
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal dblWidth As Double = 0, Optional ByVal dblTop As Double = 0, Optional ByVal Hight As Double = 0)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim ofolder As SAPbouiCOM.Folder
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 2
                    .Top = objOldItem.Top

                ElseIf position.ToUpper = "DOWN" Then
                    If ItemUID = "edWork" Then
                        .Left = objOldItem.Left + 40
                    Else
                        .Left = objOldItem.Left
                    End If
                    .Top = objOldItem.Top + objOldItem.Height + 2

                    .Width = objOldItem.Width
                    .Height = objOldItem.Height
                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            .LinkTo = linkedUID
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            objNewItem.Width = objOldItem.Width '+ 50
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            ofolder = objNewItem.Specific
            ofolder.Caption = strCaption
            ofolder.GroupWith(linkedUID)
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption

        End If
        If dblWidth <> 0 Then
            objNewItem.Width = dblWidth
        End If

        If dblTop <> 0 Then
            objNewItem.Top = objNewItem.Top + dblTop
        End If
        If Hight <> 0 Then
            objNewItem.Height = objNewItem.Height + Hight
        End If
    End Sub
#End Region
#Region "GetDocumentQuantity"
    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strTemp1 As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select CurrCode  from OCRN")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strQuantity = strQuantity.Replace(oRec.Fields.Item(0).Value, "")
            oRec.MoveNext()
        Next
        strTemp1 = strQuantity
        strTemp = CompanyDecimalSeprator
        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", ",")
        End If
        If strQuantity = "" Then
            Return 0
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = Convert.ToDouble(strTemp1)
        End Try

        Return dblQuant
    End Function
#End Region
#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function

    Public Function getMaxBatchCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS varchar)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

#End Region

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function


    
End Class
