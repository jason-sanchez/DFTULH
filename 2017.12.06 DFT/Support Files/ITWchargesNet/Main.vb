Imports System
Imports System.IO
Imports System.Configuration.Configuration


Module Main

    Sub Main()
        Dim vsDbLoc As String = System.Configuration.ConfigurationManager.AppSettings("dbLocation")
        Dim vsDbCategory As String = System.Configuration.ConfigurationManager.AppSettings("dbCategory")
        Dim vsOutputDir As String = System.Configuration.ConfigurationManager.AppSettings("outputDir")
        Dim vsLogFile As String = System.Configuration.ConfigurationManager.AppSettings("logFile")
        Dim LogStream As StreamWriter
        Dim aPatStatus, aBatchFac As String()
        Dim vsOutputTestFile As String = vsOutputDir + "test_itwCharges.txt"
        Dim OutputTestStream As StreamWriter
        Dim vsOutputFile As String = vsOutputDir + "itwCharges.txt"
        Dim OutputStream As StreamWriter
        Dim i, TypeOutput As Integer
        Dim sql As String
        Dim ds As DataSet = New DataSet
        Dim viRecCount As Integer = 0
        Dim dr, dr2 As DataRow
        Dim vsDepartmentID As String = ""
        Dim viTransmissionBatch As Integer
        Dim viBatchID As Integer
        Dim TempOutput As String = ""

        ' Establish connection to the log file
        Try
            LogStream = New StreamWriter(vsLogFile, True)
        Catch ex As Exception
            Exit Sub
        End Try

        ' Test connection to database
        Try
            dbAccess.TestConnection()
        Catch ex As Exception
            LogStream.WriteLine(Now.ToString() & " Error connecting to database" & vbCrLf & ex.Message)
            LogStream.Close()
            Exit Sub
        End Try

        'Begin the process -- first check for CPT Exclusions that were missed
        LogStream.WriteLine(vbCrLf & "***************************************************************")
        LogStream.WriteLine(Now.ToString() & " Checking for CPT Exclusions")

        ' This query compares all unbatched charges against other charges on the same day
        ' for the same patient to see if an exclusion has been missed (based on the order
        ' in which the charges were posted). To accomplish this, we have joined one instance
        ' of the charge detail table to itself (ignoring the comparison of a record to itself)
        ' and have linked each instance to its [007snPlan] counterpart to get the CPT associated
        ' with the charge, and have incorporated the [130cptExclusions] table to see if the two
        ' CPT codes conflict. The query results are only those records that need to be modified,
        ' along with the excludeAction that needs to be taken.

        ' 10/01/2007 Matt - This query did not take credits into consideration. It was discovered
        ' when a Re-Eval was accidentally charged (and thus credited), this query was blocking
        ' the Eval that should have been billed originally because of that Re-Eval. To fix, we
        ' are now looking for records based on the net quantity -- serviceQty + credits. This way
        ' if something was not supposed to be bill (ie- credited), it will be ignored.

        'sql = "select e.epnum, d.recordID as exRecordID, d.panum as exPANum, s.cpt as exCPT, "
        'sql = sql & "convert(nvarchar,d.serviceDate,101) as exServiceDate, d2.panum as PANum, "
        'sql = sql & "convert(nvarchar,d2.serviceDate,101) as serviceDate, s2.cpt as CPT, d2.recordID as recordID, "
        'sql = sql & "ex.excludeAction, ex.mutuallyExclusive, d.procMod, d.procMod2, s.[id] as snPlanID "
        'sql = sql & "from [075chargeDt] d "
        'sql = sql & "left join [075chargeDT] d2 on d.panum = d2.panum "
        'sql = sql & " and convert(nvarchar,d.serviceDate,101) = convert(nvarchar,d2.serviceDate,101) "
        'sql = sql & " and d.recordid <> d2.recordid "
        'sql = sql & " and d2.serviceQty > 0 "
        'sql = sql & " and d2.serviceCode <> '20001095' "
        'sql = sql & "inner join [007snPlan] s on d.snhdid = s.snhdid "
        'sql = sql & " and d.serviceCode = s.serviceCode "
        'sql = sql & "left join [007snPlan] s2 on d2.snhdid = s2.snhdid "
        'sql = sql & " and d2.serviceCode = s2.serviceCode "
        'sql = sql & "inner join [001episode] e on d.panum = e.panum "
        'sql = sql & "inner join [130cptExclusion] ex on left(s.cpt, 5) = ex.cptExclude "
        'sql = sql & " and left(s2.cpt, 5) = ex.cpt "
        'sql = sql & " and ex.inactive = 0 "
        '' 4/25/2007 - Only look for outpatients, per Max
        'sql = sql & "where e.status in ('OA') " ' removed: 'IA', 'ID', 'IC'
        'sql = sql & "and e.intakeFacility in (200,300) "
        'sql = sql & "and e.testCase = 0 "
        'sql = sql & "and d.serviceQty > 0 "
        'sql = sql & "and (d.batchID = 0 or d.batchID is null) "
        'sql = sql & "and d.panum > '' "
        '' 20001095 = Discharge service code
        'sql = sql & "and d.serviceCode <> '20001095' "
        'sql = sql & "and (ex.excludeAction = 0 "
        'sql = sql & "or (ex.excludeAction = 1 and d.procMod is null) "
        'sql = sql & "or (ex.excludeAction = 1 and d.procMod2 is null and d.procMod <> '59')) "
        'sql = sql & "order by d.panum, s.cpt, exServiceDate "

        sql = "select e.epnum, d.recordID as exRecordID, d.panum as exPANum, s.cpt as exCPT, "
        sql = sql & "convert(nvarchar,d.serviceDate,101) as exServiceDate, d2.panum as PANum, "
        sql = sql & "convert(nvarchar,d2.serviceDate,101) as serviceDate, s2.cpt as CPT, d2.recordID as recordID, "
        sql = sql & "ex.excludeAction, ex.mutuallyExclusive, d.procMod, d.procMod2, s.[id] as snPlanID "
        sql = sql & "from [075chargeDt] d "
        sql = sql & "left join [075chargeDT] d2 on d.panum = d2.panum "
        sql = sql & "and convert(nvarchar,d.serviceDate,101) = convert(nvarchar,d2.serviceDate,101) "
        sql = sql & "and d.recordid <> d2.recordid "
        sql = sql & "and (d2.serviceQty + (-1 * ISNULL(d2.credits, 0))) > 0 "
        sql = sql & "and d2.serviceCode <> '20001095' "
        sql = sql & "inner join [007snPlan] s on d.snhdid = s.snhdid "
        sql = sql & "and d.serviceCode = s.serviceCode "
        sql = sql & "left join [007snPlan] s2 on d2.snhdid = s2.snhdid "
        sql = sql & "and d2.serviceCode = s2.serviceCode "
        sql = sql & "inner join [001episode] e on d.panum = e.panum "
        sql = sql & "inner join [113payor] p on e.primaryIns = p.planCode "
        sql = sql & "inner join [130cptExclusion] ex on left(s.cpt, 5) = ex.cptExclude "
        sql = sql & "and left(s2.cpt, 5) = ex.cpt "
        sql = sql & "and ex.inactive = 0 "
        ' 4/25/2007 - Only look for outpatients, per Max
        sql = sql & "where e.status in ('OA') " ' removed: 'IA', 'ID', 'IC'
        ' 06/10/2008 Matt - Only look for OP Medicare now
        sql = sql & "and (p.medicare = 1 "
        ' 08/26/2008 Matt - Also look for OP Medicaid
        sql = sql & "or p.medicaid = 1) "
        sql = sql & "and e.intakeFacility in (200,300) "
        sql = sql & "and e.testCase = 0 "
        sql = sql & "and (d.serviceQty + (-1 * ISNULL(d.credits, 0))) > 0 "
        sql = sql & "and (d.batchID = 0 or d.batchID is null) "
        sql = sql & "and d.panum > '' "
        ' 20001095 = Discharge service code
        sql = sql & "and d.serviceCode <> '20001095' "
        sql = sql & "and (ex.excludeAction = 0 "
        sql = sql & "or (ex.excludeAction = 1 and d.procMod is null) "
        sql = sql & "or (ex.excludeAction = 1 and d.procMod2 is null and d.procMod <> '59')) "
        sql = sql & "order by d.panum, s.cpt, exServiceDate "

        Try
            dbAccess.fillDataSetTable(ds, sql, "CPTExclusions")
        Catch ex As Exception
            LogStream.WriteLine(Now.ToString() & " Error querying for missing modifiers" & vbCrLf & ex.Message)
            LogStream.Close()
            Exit Sub
        End Try

        For Each dr In ds.Tables("CPTExclusions").Rows
            ' Try to apply the necessary exclusions, but don't hold up the batching process if unable to
            Try
                TempOutput = ""

                ' The action taken here depends on the excludeAction associated with the row
                Select Case (Convert.ToInt32(dr.Item("excludeAction").ToString()))
                    Case 0
                        ' This charge should not be billed at all, so set the quantity to 0 and give it
                        ' a bogus batch of -50
                        sql = "update [075chargeDt] set "
                        sql &= "serviceQty = 0, "
                        sql &= "batchID = -50 "
                        sql &= "where recordID = " & dr.Item("exRecordID").ToString() & " "

                        ' Only continue if successful
                        If dbAccess.executeNonQuery(sql) Then
                            ' Log that the action was successful
                            TempOutput = Chr(9) & "ChargeDT " & dr.Item("exRecordID").ToString() & " - Charge not allowed"

                            ' Update [007snPlan], set the quantity to 0 and update the billingAction
                            sql = "update [007snPlan] set "
                            sql &= "quantity = 0, "
                            sql &= "BillingAction = " & dr.Item("excludeAction").ToString() & " "
                            sql &= "where [id] = " & dr.Item("snPlanID").ToString() & " "
                            dbAccess.executeNonQuery(sql)

                            LogStream.WriteLine(TempOutput & " (snPlan " & dr.Item("snPlanID").ToString() & " updated as well)")
                        End If
                    Case 1
                        ' Update [075chargeDt] to add the '59' modifier
                        sql = "update [075chargeDt] set "
                        If IsDBNull(dr.Item("procMod")) Then
                            ' Add the '59' to the procMod field
                            sql &= "procMod = '59' "
                        Else
                            ' Add the '59' to the procMod2 field
                            sql &= "procMod2 = '59' "
                        End If
                        sql &= "where recordID = " & dr.Item("exRecordID").ToString() & " "

                        ' Only update the [007snPlan] if the [075chargeDt] record is updated
                        ' (barring other errors, of course)
                        If dbAccess.executeNonQuery(sql) Then
                            ' Log that the action was successful
                            TempOutput = Chr(9) & "ChargeDT " & dr.Item("exRecordID").ToString() & " - Added modifier 59"

                            ' update [007snPlan] to set the billingAction (which is the same as the excludeAction)
                            sql = "update [007snPlan] set "
                            sql &= "BillingAction = " & dr.Item("excludeAction").ToString() & " "
                            sql &= "where [id] = " & dr.Item("snPlanID").ToString() & " "
                            dbAccess.executeNonQuery(sql)

                            LogStream.WriteLine(TempOutput & " (snPlan " & dr.Item("snPlanID").ToString() & " updated as well)")
                        End If
                    Case Else
                        ' Unknown exclude action -- just log it
                        LogStream.WriteLine(Chr(9) & "ChargeDT " & dr.Item("exRecordID").ToString() & " - Unknown exclude action - " & dr.Item("excludeAction"))
                End Select
            Catch ex As Exception
                Try
                    ' Log what was done
                    If TempOutput > "" Then
                        LogStream.WriteLine(TempOutput)
                    End If

                    ' Try to indicate which charge record the process failed on
                    LogStream.WriteLine(Chr(9) & "ChargeDT " & dr.Item("exRecordID").ToString() & " - UNABLE TO PROCESS EXCLUSION (" & ex.Message & ")")
                Catch ex2 As Exception
                    ' As a last resort, at least say that the process failed and give the specific errors
                    LogStream.WriteLine(Chr(9) & "UNABLE TO PROCESS EXCLUSION [" & ex.Message & "] [" & ex2.Message & "]")
                End Try
            End Try
        Next

        ' Create arrays to control how the charges are grouped into batches
        aPatStatus = New String() {"OA", "IA','ID','IC", "OA", "IA','ID','IC"}
        aBatchFac = New String() {"200", "200", "300", "300"}

        ' Establish connections to the output files
        Try
            OutputTestStream = New StreamWriter(vsOutputTestFile, False)
            OutputStream = New StreamWriter(vsOutputFile, True)
        Catch ex As Exception
            LogStream.WriteLine(Now.ToString() + " Error creating Output Files")
            LogStream.Close()
            Exit Sub
        End Try

        ' Batch the charges by looping through the arrays
        'Loop through Frazier OP, Frazier IP, SIRH OP & SIRH IP
        For i = 0 To 3
            Try
                LogStream.WriteLine("====================" & vbCrLf & Now.ToString() + " Selecting data for " + Replace(aPatStatus(i), "'", "") + " " + aBatchFac(i))

                'Main query to get all charges
                sql = "select d.recordId, d.paNum, d.serviceDate, d.serviceCode, d.serviceQty, "
                sql = sql & "d.procMod, d.procMod2, depart.billingId, depart.billingStatus "
                sql = sql & "from [075chargeDt] as d "
                sql = sql & "inner join [001episode] as e on d.panum = e.panum "
                sql = sql & "inner join [130HospSvc] as h on (e.patient_Type = h.patientType "
                sql = sql & "   and e.intakeFacility = h.intakeFacility "
                sql = sql & "   and e.hservice = h.hospSvc) "
                sql = sql & "inner join [130department] as depart on (h.department = depart.department "
                sql = sql & "   and h.intakeFacility = depart.intakeFacility "
                sql = sql & "   and h.patientType = depart.patientType) "
                ' 2/8/2011 Matt - #5159 - Only pull charges with a billing agent of 0.
                '   This is to filter out Psych right now, but might be used to filter
                '   other charges in the future.
                sql = sql & "inner join [130Service] as serv on d.serviceCode = serv.servCode "
                sql = sql & "where (d.batchID = 0 or d.batchID is null) "
                sql = sql & "and e.status in ('" & aPatStatus(i) & "') "
                sql = sql & "and e.intakeFacility = " & aBatchFac(i) & " "
                sql = sql & "and e.testCase = 0 "
                ' 2/8/2011 Matt - #5159 - Only pull charges with a billing agent of 0.
                sql = sql & "and serv.billingAgent = 0 "
                sql = sql & "order by depart.billingId, d.recordId, d.paNum, d.serviceDate, "
                sql = sql & "d.serviceCode, d.serviceQty, d.procMod "

                If ds.Tables.Contains("Charges") Then
                    ds.Tables("Charges").Clear()
                End If

                dbAccess.fillDataSetTable(ds, sql, "Charges")

                'query to get next transmission batch
                sql = "select top 1 transmissionbatch "
                sql = sql & "from [075chargeHD] "
                sql = sql & "order by batchID desc"

                If ds.Tables.Contains("LastTransBatch") Then
                    ds.Tables("LastTransBatch").Clear()
                End If

                dbAccess.fillDataSetTable(ds, sql, "LastTransBatch")

                If ds.Tables("LastTransBatch").Rows.Count > 0 Then
                    If Not IsDBNull(ds.Tables("LastTransBatch").Rows(0).Item("transmissionBatch")) Then
                        viTransmissionBatch = ds.Tables("LastTransBatch").Rows(0).Item("transmissionBatch")
                    Else
                        viTransmissionBatch = 0 'Start at 0 since it will increment right away
                    End If
                Else
                    viTransmissionBatch = 0 'Start at 0 since it will increment right away
                End If

                LogStream.WriteLine(Now.ToString() + " Creating Batches")

                For Each dr In ds.Tables("Charges").Rows
                    If Not IsDBNull(dr.Item("billingID")) Then
                        If vsDepartmentID <> dr.Item("billingID") Then
                            If viRecCount > 0 Then
                                sql = "update [075chargeHD] set numRecords = " & viRecCount.ToString
                                sql = sql & ", numTransactions = " & viRecCount.ToString
                                sql = sql & "where batchID = " & viBatchID.ToString

                                dbAccess.executeNonQuery(sql)
                            End If

                            viTransmissionBatch += 1

                            If viTransmissionBatch > 999 Then
                                viTransmissionBatch = 1
                            End If

                            'insert next batch header
                            sql = "set nocount on "
                            sql = sql & "insert into [075chargeHD] (transmissionBatch, patientStatus, batchFac, "
                            sql = sql & "BatchDate, departmentId, billingStatus, numRecords, numTransactions)"
                            sql = sql & "values("
                            sql = sql & viTransmissionBatch.ToString & ", "
                            sql = sql & "'" & Replace(aPatStatus(i), "'", "''") & "', "
                            sql = sql & aBatchFac(i) & ", "
                            sql = sql & "'" & Now.ToString & "', "
                            sql = sql & "'" & dr.Item("billingId").ToString & "', "
                            sql = sql & dr.Item("billingStatus").ToString & ", "
                            sql = sql & "0, "
                            sql = sql & "0) "
                            sql = sql & "select @@Identity as BatchID "
                            sql = sql & "set nocount off "

                            If ds.Tables.Contains("CurrentBatch") Then
                                ds.Tables("CurrentBatch").Clear()
                            End If

                            dbAccess.fillDataSetTable(ds, sql, "CurrentBatch")

                            viBatchID = ds.Tables("CurrentBatch").Rows(0).Item("BatchID")

                            vsDepartmentID = dr.Item("billingID")
                            viRecCount = 0
                        End If

                        sql = "update [075chargeDT] set batchID = " & viBatchID.ToString & " "
                        sql = sql & "where recordID = " & dr.Item("recordID").ToString

                        dbAccess.executeNonQuery(sql)

                        viRecCount += 1
                    End If
                Next

                ' Cleanup
                If viRecCount > 0 Then
                    sql = "update [075chargeHD] set numRecords = " & viRecCount.ToString
                    sql = sql & ", numTransactions = " & viRecCount.ToString
                    sql = sql & "where batchID = " & viBatchID.ToString

                    dbAccess.executeNonQuery(sql)

                    viRecCount = 0
                End If
            Catch ex As Exception
                LogStream.WriteLine(Now.ToString() + " ***ERROR - Batch Creation***")
                LogStream.WriteLine("  " + ex.Message)
                Exit Sub
            End Try

            '****SECTION 2 - CREATE TEXT FILES
            'Loop to create output file (billing status 2) & test file (billing status 1)
            For TypeOutput = 0 To 1

                'Write Charge Header Records to Files
                sql = "select * from [075chargeHD] "
                sql = sql & "where transcreated = 0 "
                sql = sql & "and billingStatus = " & (TypeOutput + 1).ToString & " "
                sql = sql & "order by transmissionBatch "

                If ds.Tables.Contains("Batches") Then
                    ds.Tables("Batches").Clear()
                End If

                dbAccess.fillDataSetTable(ds, sql, "Batches")

                For Each dr In ds.Tables("Batches").Rows

                    'create header record
                    TempOutput = "D " 'RecordType
                    TempOutput &= ZeroLeft(dr.Item("transmissionBatch").ToString, 3) 'continuousBatchNumber
                    'TempOutput &= StringOfChar(" ", 4) 'Filler 1
                    TempOutput &= "03" 'Filler 2
                    TempOutput &= StringOfChar(" ", 8) 'Filler 3
                    Select Case (i + 1)
                        Case 1, 2
                            TempOutput &= "H080" 'Facility
                        Case 3, 4
                            TempOutput &= "M080" 'Facility
                    End Select
                    TempOutput &= Format(Now, "MMddyy") 'BatchDate
                    TempOutput &= StringOfChar(" ", 4) 'Filler 4
                    TempOutput &= dr.Item("departmentID").ToString 'BatchIdentifier
                    TempOutput &= StringOfChar(" ", 3) 'Filler 5
                    Select Case (i + 1)
                        Case 1, 2
                            TempOutput &= "FR7" 'FixedBatchIdentifier
                        Case 3, 4
                            TempOutput &= "RI7" 'FixedBatchIdentifier
                    End Select
                    TempOutput &= ZeroLeft(dr.Item("transmissionBatch").ToString, 3) 'continuousBatchNumber
                    TempOutput &= StringOfChar(" ", 38) 'Filler 6

                    Select Case TypeOutput + 1
                        Case 1 'Write to Test File
                            OutputTestStream.WriteLine(TempOutput)
                        Case 2 'Write to Real Output File
                            OutputStream.WriteLine(TempOutput)
                    End Select


                    'Write Charge Detail Records to Files
                    sql = "select * from [075chargeDT] "
                    sql = sql & "where BatchID = " & dr.Item("BatchID").ToString & " "
                    sql = sql & "order by serviceDate, recordID "

                    If ds.Tables.Contains("Detail") Then
                        ds.Tables("Detail").Clear()
                    End If

                    dbAccess.fillDataSetTable(ds, sql, "Detail")

                    For Each dr2 In ds.Tables("Detail").Rows

                        'create detail record
                        If Not IsDBNull(dr2.Item("serviceQty")) AndAlso Convert.ToInt32(dr2.Item("serviceQty")) > 0 Then
                            TempOutput = "48" 'RecordType
                        Else
                            TempOutput = "49" 'RecordType
                        End If

                        TempOutput &= ZeroLeft(dr2.Item("panum").ToString, 12) 'Panum
                        TempOutput &= StringOfChar(" ", 1) 'Filler 1
                        TempOutput &= Format(dr2.Item("serviceDate"), "MMddyy") 'ServiceDate
                        TempOutput &= PadRight(dr2.Item("serviceCode"), 8) 'ServiceCode
                        TempOutput &= StringOfChar(" ", 2) 'Filler 2
                        If Not IsDBNull(dr2.Item("serviceQty")) AndAlso Convert.ToInt32(dr2.Item("serviceQty")) > 0 Then
                            TempOutput &= ZeroLeft(dr2.Item("serviceQty").ToString, 5) 'Quantity
                        Else
                            TempOutput &= ZeroLeft(Replace(dr2.Item("serviceQty").ToString, "-", ""), 5) 'Quantity
                        End If
                        TempOutput &= StringOfChar(" ", 14) 'Filler 3
                        If IsDBNull(dr2.Item("procMod2")) Then
                            TempOutput &= StringOfChar(" ", 2) 'Modifier2
                        Else
                            TempOutput &= dr2.Item("procMod2").ToString 'Modifier2
                        End If
                        TempOutput &= StringOfChar(" ", 2) 'Modifier3
                        TempOutput &= StringOfChar(" ", 16) 'Filler 3a
                        'If IsDBNull(dr2.Item("procMod")) Then
                        'TempOutput &= StringOfChar(" ", 2) 'Modifier1
                        'Else
                        TempOutput &= dr2.Item("procMod").ToString 'Modifier1
                        'End If
                        TempOutput &= StringOfChar(" ", 7) 'Filler 4

                        Select Case TypeOutput + 1
                            Case 1 'Write to Test File
                                OutputTestStream.WriteLine(TempOutput)
                            Case 2 'Write to Real Output File
                                OutputStream.WriteLine(TempOutput)
                        End Select

                    Next 'Loop ChargeDT Records to write to output files

                    If Convert.ToInt32(dr.Item("transmissionBatch").ToString()) > 130 Then
                        ' break here
                        TempOutput = ""
                    End If

                    'Create footer
                    TempOutput = "98" 'RecordType
                    TempOutput &= StringOfChar(" ", 3) 'Filler 1
                    TempOutput &= ZeroLeft(dr.Item("numRecords").ToString, 3) 'ChargeRecordCount
                    TempOutput &= ZeroLeft(Convert.ToInt32(dr.Item("numRecords").ToString) + 2, 3) 'RecordCount
                    TempOutput &= StringOfChar("0", 15) 'Filler 2

                    Select Case TypeOutput + 1
                        Case 1 'Write to Test File
                            OutputTestStream.WriteLine(TempOutput)
                        Case 2 'Write to Real Output File
                            OutputStream.WriteLine(TempOutput)
                    End Select

                    sql = "update [075chargeHD] set transCreated = 1 "
                    sql = sql & "where batchID = " & dr.Item("batchID").ToString
                    dbAccess.executeNonQuery(sql)

                Next 'Loop ChargeHD Records to write to output files
            Next 'Loop to create output file (billing status 2) & test file (billing status 1)
        Next 'Loop through Frazier OP, Frazier IP, SIRH OP & SIRH IP

        LogStream.WriteLine("====================" & vbCrLf & Now.ToString() + " Process completed successfully")
        LogStream.Close()
        OutputTestStream.Close()
        OutputStream.Close()

    End Sub
    Private Function ZeroLeft(ByVal input As String, ByVal FinalLength As Integer) As String
        While input.Length < FinalLength
            input = "0" & input
        End While
        Return input
    End Function
    Private Function StringOfChar(ByVal theChar As String, ByVal Length As Integer) As String
        Dim i As Integer
        Dim output As String = ""

        For i = 0 To Length - 1
            output &= theChar
        Next

        Return output
    End Function
    Private Function PadRight(ByVal input As String, ByVal FinalLength As Integer) As String
        While input.Length < FinalLength
            input &= " "
        End While
        Return input
    End Function
End Module
