Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try

            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OADM" Or strTab = "ORDR" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OCRN" Or strTab = "OITM" Or strTab = "RDR1") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"


    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If
            objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    '  MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
            Else


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub
    Public Function UDOPump(ByVal strUDO As String, _
                       ByVal strDesc As String, _
                           ByVal strTable As String, _
                               ByVal intFind As Integer, _
                                   Optional ByVal strCode As String = "", _
                                       Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_PumpCode"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_PumpCode"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_PumpDesc"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_PumpDesc"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Active"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Active"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    ' oUserObjects.FindColumns.Add()
                    'oUserObjects.FindColumns.SetCurrentLine(1)
                    'oUserObjects.FindColumns.ColumnAlias = strName
                    'oUserObjects.FindColumns.Add()
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function


    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try

            'oProgressBar = oApplication.SBO_Application.StatusBar.CreateProgressBar("Initializing Database...", 8, False)
            'oProgressBar.Value = 0
            'oProgressBar.Text = "Initializing Database... "


            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            ' AddFields("OCRN", "AC", "Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddTables("Z_ADCS", "Additional Cost Setup", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_ADCS", "Z_WhoutPump", "WithOut Pump Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_ADCS", "Z_WaterProof", "Water Proof Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_ADCS", "Z_DefTemp", "Default Temperature ", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_ADCS", "Z_Temp", "Temperature Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            addField("Z_ADCS", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_ADCS", "Z_FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_ADCS", "Z_ToDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddTables("Z_ADC1", "Additional Cost Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_ADC1", "Z_ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_ADC1", "Z_UnitPrice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            addField("Z_ADC1", "Z_WhoutPump", "WithOut Pump", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("Z_ADC1", "Z_WaterProof", "Water Proof", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_ADC1", "Z_DefTemp", "Default Temperature ", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_ADC1", "Z_Temp", "Actual Temperature", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_ADC1", "Z_ActualPrice", "Actual Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_ADC1", "Z_RefCode", "Additional Cost Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ADC1", "Z_WhPumpPrice", "WithOut Pump Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_ADC1", "Z_WatProofPrice", "Water Proof Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_ADC1", "Z_TempValue", "Temperature Value", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)

            AddTables("Z_ADC2", "Additional Cost Item Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_ADC2", "Z_AddRef", "Additional Cost Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ADC2", "Z_ItemCode", "Additional Cost Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ADC2", "Z_ItemName", "Additional Cost Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_ADC2", "Z_Price", "Additional Cost ItemPrice", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_ADC2", "Z_Quantity", "Additional Cost Item Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)

            addField("OITM", "Z_ItemType", "Additional Cost Type Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("OITM", "Z_ItemCost", "Additional Cost Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("OITM", "Z_DefQty", "Default Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)


            AddFields("ORDR", "Z_DupEntry", "Number of Duplicate Entry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ORDR", "Z_SourceDoc", "Source Document Entry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ORDR", "Z_StartingDate", "DuplicateEntry Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date)


            AddTables("DORDR", "Duplicate Sales Order", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("DORDR", "Z_DupEntry", "Number of Duplicate Entry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("DORDR", "Z_SourceDoc", "Source Document Entry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("DORDR", "Z_StartingDate", "Dup.Entry  Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("DORDR", "Z_SourceDocNum", "Source Document Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("DORDR", "Z_Winch", "Winch", SAPbobsCOM.BoFieldTypes.db_Numeric)

            addField("ITT1", "Z_Type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "R,A", "Regular,Additional Cost", "R")


            AddFields("RDR1", "Z_AddRef", "Additional Cost Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("RDR1", "Z_WhoutPump", "WithOut Pump", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("RDR1", "Z_WaterProof", "Water Proof", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("RDR1", "Z_Temp", "Actual Temperature", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

          
            AddFields("ORDR", "Z_PumpCode", "Pump Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("ORDR", "Z_PumpName", "Pump Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("ORDR", "Z_FromTime", "From Time", SAPbobsCOM.BoFieldTypes.db_Date, , , SAPbobsCOM.BoFldSubTypes.st_Time)
            AddFields("ORDR", "Z_ToTime", "To Time", SAPbobsCOM.BoFieldTypes.db_Date, , , SAPbobsCOM.BoFldSubTypes.st_Time)
            AddFields("ORDR", "Z_FromDate", "Delivery from date and Time", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("ORDR", "Z_ToDate ", "Delivery To date and Time", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

          
            AddTables("Z_OPUMP", "Pump Master Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OPUMP", "Z_PumpCode", "Pump Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPUMP", "Z_PumpDesc", "Pump Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OPUMP", "Z_FromDate", "InActive From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPUMP", "Z_ToDate", "InActive ToDate", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPUMP", "Z_InActive", "Active Status", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("@Z_OPUMP", "Z_Active", "Pump Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            AddTables("Z_OACAD", "Add.Cost Allocation Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OACAD", "Z_DocNum", "Sales Order Doc.Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OACAD", "Z_DocEntry", "Sales Order Doc.Entry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OACAD", "Z_DocDate", "Sales Order Doc.Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OACAD", "Z_DelDate", "Sales Order Del.Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OACAD", "Z_PumpCode", "Pump Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OACAD", "Z_PumpName", "Pump Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OACAD", "Z_FromHour", "From Hour", SAPbobsCOM.BoFieldTypes.db_Date, , , SAPbobsCOM.BoFldSubTypes.st_Time)
            AddFields("Z_OACAD", "Z_ToHour", "To Hour", SAPbobsCOM.BoFieldTypes.db_Date, , , SAPbobsCOM.BoFldSubTypes.st_Time)

            AddFields("RDR1", "Z_LOCATION", "Location", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("RDR1", "Z_TYPE", "TYPE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)


            oApplication.Company.StartTransaction()

            '---- User Defined Object's
            CreateUDO()

            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If

        Catch ex As Exception
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Throw ex
        Finally
            'oProgressBar.Stop()
            ' System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try
            AddUDO("Z_ADCS", "Additional Cost Setup", "Z_ADCS", "DocEntry", , , SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
