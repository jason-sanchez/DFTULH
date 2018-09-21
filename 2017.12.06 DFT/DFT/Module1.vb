Imports System
Imports System.IO
Imports System.Collections
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports Ionic.Zip
'20130417 - added filters to remove unwanted charges
'20130103 - changed ITW connection string. Old: 10.48.10.246,1433.  New: 10.48.242.249,1433.
'20140319 - version for wave 3 testing on Sysfeed5
'20140325 - mod for cscsysfeed5 (10.48.64.5)
'20140407 terminate FT1 with cr instead of crlf. Use STAR Region in generateMSH routine.
'20140407- change credit from CR to CD per phoncon 20140407 (Noreen Marcum and Bill Glover)
'20140407 - update some MSH fields per tri-map received 4/7/2014.
'20140430 - pad left mrnum to 9 zeros in PID function.
'20140515 - use SIMCode and SIMDepartment in FT1 generator.
'20140817 - mods for W3 Production.

'20141208 = removed toString reference particularly in the PID generator.

'20150413 - VS2013 version
Module Module1
    '20121022 - generate DFT files

    '20121119 - Added control code wrapper around text file. These are the same codes that were used
    'with the acknowledgement section of the HL7 Receiver program.

    '20121120 - Added patient type in PV1_2.

    'Public objIniFile As New INIFile("d:\w3Production\HL7Mapper.ini") '20140817 'Prod
    Public objIniFile As New INIFile("c:\W3Feeds\HL7Mapper.ini") '20140817 'Test
    'Public objIniFile As New INIFile("C:\KY1 Test Environment\HL7Mapper.ini") 'Local

    Dim sql As String = ""
    Dim strOutputDirectory As String = ""
    Dim strOutputSubDirectory As String = ""
    Dim gblStrPostDate As String = ""
    Dim strHL7Output As String = ""

    'Public conIniFile As New INIFile("d:\W3Production\KY1ConnProd.ini") '20140805 Prod
    'Public conIniFile As New INIFile("C:\KY1 Test Environment\KY1ConnDev.ini") 'Local
    Public conIniFile As New INIFile("C:\W3Feeds\KY1ConnTest.ini") 'Test

    Dim connString = conIniFile.GetString("Strings", "DFT", "(none)")
    'Dim connString = "server=10.48.242.249,1433;database=ITW;uid=sysmax;pwd=Condor!" '20140817 'PROD
    'Dim connString = "server=10.48.64.5\sqlexpress;database=ITWTest;uid=sysmax;pwd=Condor!" 'TEST
    'Dim connString = "server=192.168.55.12\sqlexpress;database=ITW;uid=sysmax;pwd=sysmax" '12
    Dim strPANUM As String = ""

    Dim gblFT1Count As Integer = 0
    Public Sub SendMessage(ByRef strMsg As String)
        Dim file As System.IO.StreamWriter
        'file = My.Computer.FileSystem.OpenTextFileWriter("d:\DFTLog\DFTlog.txt", True) '20140817 Prod
        file = My.Computer.FileSystem.OpenTextFileWriter("c:\DFTLog\ULH\DFTlog.txt", True) '20140817 Test
        'file = My.Computer.FileSystem.OpenTextFileWriter("C:\KY1 Test Environment\DFTLog\DFTlog.txt", True) '20140817 Local
        file.WriteLine(strMsg)
        file.Close()

    End Sub
    Sub Main()
        strOutputDirectory = objIniFile.GetString("DFTULH", "DFToutputdirectory", "(none)") 'c:\feedsW3\DFT\
        Dim DFTConnection As New SqlConnection(connString)
        Dim DFTCommand As New SqlCommand
        DFTCommand.Connection = DFTConnection
        Dim dr_DFT As SqlDataReader
        
        Try
            'Find all pa numbers in the 075ChargeDT table that have the HL7Generated field set to null.
            'This means tha the record has not been entered into an FT1 segnment yet.
            'sql = "select distinct panum from [075ChargeDt] "
            'sql = sql & "where(hl7Generated Is null) "

            '20130417 - add filters
            sql = "select distinct c.panum from [075ChargeDt] c, [001episode] e  "
            sql = sql & "where(c.hl7Generated Is null) "
            sql = sql & "and c.panum = e.panum "
            sql = sql & "and e.Status not in ('OP','IP','OX','IX') "
            sql = sql & "and e.Testcase = 0 "
            '20180920 - ULH only
            sql = sql & "and e.intakefacility IN (400)"

            DFTCommand.CommandText = sql
            DFTConnection.Open()
            dr_DFT = DFTCommand.ExecuteReader()

            If dr_DFT.HasRows Then

                Do While dr_DFT.Read()
                    strPANUM = dr_DFT.Item("panum")
                    GenerateMSH(strPANUM) '20140523
                    GeneratePID(strPANUM) '20140523
                    GeneratePV1(strPANUM) '20140523
                    GenerateFT1(strPANUM)

                    If gblFT1Count > 0 Then '20140523
                        CreateOutputFile(strHL7Output) '20140523
                    End If '20140523
                Loop

            End If



            '20130226 - compress the backup files
            compressfiles(strOutputDirectory & "\backup")
            dr_DFT.Close()
            DFTConnection.Close()

            '20130228 - handle the left over records with billing agent = 1
            handleBillingAgent()

            '20130417 - handle filters. Testcase and deleted episodes.
            handleFilters()

        Catch ex As Exception
            'SendMessage("Main Processing Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            'SendMessage("DFT Generation Complete" & vbCrLf)
        End Try

    End Sub
    Public Sub GenerateFT1(ByRef panum As String)
        '201217 - changed sql statement to extract the service code for the correct intake facility
        '200 H80 Frazier; 300 M80 SIRH
        'this code will not show a record that is not in the 130service table
        Dim sql As String = ""
        Dim sqlUpdate As String = ""
        Dim DFTConnection As New SqlConnection(connString)
        Dim DFTCommand As New SqlCommand
        DFTCommand.Connection = DFTConnection
        Dim dr_DFT As SqlDataReader



        Dim updateConnection As New SqlConnection(connString)
        Dim UpdateCommand As New SqlCommand
        UpdateCommand.Connection = updateConnection

        Dim FT1Counter As Integer = 0
        Dim ft1Segment As String = ""
        Dim intTransactionType As Integer = 0
        Try


            'sql = "SELECT [075ChargeDt].batchID, [075ChargeDt].recordID, [075ChargeDt].serviceDate, [075ChargeDt].serviceQty, [075ChargeDt].procMod, [075ChargeDt].procMod2, "
            'sql = sql & "[130Service].description, [075ChargeDt].serviceCode, [075ChargeDt].paNum "
            'sql = sql & "FROM [001Episode] left JOIN [075ChargeDt] ON [001Episode].PANum = [075ChargeDt].paNum left JOIN "
            'sql = sql & "[130Service] ON [075ChargeDt].serviceCode = [130Service].servCode "
            'sql = sql & "WHERE     ([001Episode].PANum = '" & panum & "') "
            'sql = sql & "and [075ChargeDt].HL7Generated is null "
            'sql = sql & "order by serviceDate"

            sql = "SELECT [075ChargeDt].batchID, [075ChargeDt].recordID, [075ChargeDt].serviceDate, [075ChargeDt].serviceQty, [075ChargeDt].procMod, [075ChargeDt].procMod2, "
            sql = sql & "[130Service].description, [130Service].glKey, [130Service].SIMCode, [130Service].SIMDepartment, [075ChargeDt].serviceCode, [075ChargeDt].paNum, [001Episode].patient_type, "

            '20170629 - add phynum to DFT for ULH (if flagged)
            sql = sql & "sendPhysnumtoDFT "

            sql = sql & "FROM [001Episode] left JOIN [075ChargeDt] ON [001Episode].PANum = [075ChargeDt].paNum left JOIN "
            sql = sql & "[130Service] ON [075ChargeDt].serviceCode = [130Service].servCode "
            sql = sql & "WHERE     ([001Episode].PANum = '" & panum & "') "
            sql = sql & "and  [130Service].IntakeFacility = [001Episode].intakeFacility "
            sql = sql & "and [075ChargeDt].HL7Generated is null "

            '20130129 - don't use if billingAgent = 1 don't want to send these.
            sql = sql & "and [130Service].billingAgent = 0 "
            sql = sql & "order by serviceDate "

            DFTCommand.CommandText = sql
            DFTConnection.Open()
            dr_DFT = DFTCommand.ExecuteReader()

            If dr_DFT.HasRows Then
                Do While dr_DFT.Read()

                    'GenerateMSH(panum) '20140523
                    'GeneratePID(panum) '20140523
                    'GeneratePV1(panum) '20140523

                    'increment the counter
                    FT1Counter = FT1Counter + 1
                    'build the FI1 segments here for each record returned
                    ft1Segment = "FT1|" & FT1Counter & "|" & dr_DFT.Item("recordID") & "|"
                    ft1Segment = ft1Segment & dr_DFT.Item("batchID") & "|"
                    ft1Segment = ft1Segment & buildDate(dr_DFT.Item("serviceDate")) & "|" & gblStrPostDate & "|"

                    If CInt(dr_DFT.Item("serviceQty")) < 0 Then
                        '20140407 - change CR to CD
                        ft1Segment = ft1Segment & "CD" & "|"
                    Else
                        ft1Segment = ft1Segment & "CG" & "|"
                    End If

                    If Not IsDBNull(dr_DFT.Item("description")) Then
                        ft1Segment = ft1Segment & dr_DFT.Item("SIMCode").ToString & "^" & Trim(dr_DFT.Item("description").ToString) & "|||"
                    Else
                        ft1Segment = ft1Segment & dr_DFT.Item("SIMCode") & "|||"
                    End If
                    ft1Segment = ft1Segment & Math.Abs(CInt(dr_DFT.Item("serviceQty"))) & "|||"

                    '20121218 - add glkey to ft1_13
                    ft1Segment = ft1Segment & dr_DFT.Item("SIMDepartment").ToString & "|||||"
                    '20121218 - add patient type FT1_18
                    ft1Segment = ft1Segment & dr_DFT.Item("patient_type").ToString

                    '20170629 - add phynum to DFT for ULH (if flagged)
                    If dr_DFT.Item("sendPhysnumtoDFT") = True Then 'AndAlso dr_DFT.Item("star_region") = "T" Then

                        Dim DFTConnection2 As New SqlConnection(connString)
                        Dim DFTCommand2 As New SqlCommand
                        DFTCommand2.Connection = DFTConnection2
                        'DFTConnection.Close()
                        Dim sql1 As String = ""

                        sql1 = " SELECT TOP 1 physnum "
                        sql1 += " FROM [006potHd] p "
                        sql1 += " INNER JOIN [001episode] e on p.epnum = e.epnum "
                        sql1 += " WHERE e.panum = '" & panum & "' "
                        sql1 += " ORDER by potDate desc "

                        DFTCommand2.CommandText = sql1
                        DFTConnection2.Open()

                        'Handle null physnums
                        Dim physnumobj As Object = DFTCommand2.ExecuteScalar()
                        If IsDBNull(physnumobj) Then
                            ft1Segment = ft1Segment & "|||"
                            DFTConnection2.Close()
                        Else
                            Dim physnum As Integer = Convert.ToInt32(physnumobj)
                            DFTConnection2.Close()

                            ft1Segment = ft1Segment & "||" & physnum & "|"
                        End If

                    Else
                        ft1Segment = ft1Segment & "|||"
                    End If

                    Dim strProcMod As String = ""
                    Dim strProcMod2 As String = ""
                    If Not IsDBNull(dr_DFT.Item("procMod")) Then
                        strProcMod = dr_DFT.Item("procMod").ToString
                    End If
                    If Not IsDBNull(dr_DFT.Item("procMod2")) Then
                        strProcMod2 = dr_DFT.Item("procMod2").ToString
                    End If

                    '20170629 - add phynum to DFT for ULH (if flagged)
                    If strProcMod <> "" And strProcMod2 <> "" Then
                        ft1Segment = ft1Segment & "|||||" & dr_DFT.Item("procMod").ToString & "~" & dr_DFT.Item("procMod2").ToString
                    ElseIf strProcMod <> "" And strProcMod2 = "" Then
                        ft1Segment = ft1Segment & "|||||" & dr_DFT.Item("procMod").ToString
                    ElseIf strProcMod = "" And strProcMod2 <> "" Then
                        ft1Segment = ft1Segment & "|||||" & dr_DFT.Item("procMod2").ToString
                    ElseIf strProcMod = "" And strProcMod2 = "" Then
                        ft1Segment = ft1Segment & "|||||"
                    End If

                    'ft1Segment = ft1Segment & "||||||||" & "procmod~procmod2"
                    'add the FT1 segment to the output file

                    '20140407 - terminate line vith Cr only
                    strHL7Output = strHL7Output & ft1Segment & vbCr
                    'clear the ft1Segment sting and get ready for next segment
                    ft1Segment = ""

                    '20121105 - add  code to update the HL7Generated field with the datetime
                    sqlUpdate = "Update [075ChargeDT] "
                    sqlUpdate = sqlUpdate & "Set HL7Generated = '" & DateTime.Now & "' "
                    sqlUpdate = sqlUpdate & "WHERE recordID = " & dr_DFT.Item("recordID")

                    UpdateCommand.CommandText = sqlUpdate
                    updateConnection.Open()
                    UpdateCommand.ExecuteNonQuery()
                    updateConnection.Close()


                    'CreateOutputFile(strHL7Output) '20140523

                Loop

            End If

            dr_DFT.Close()
            DFTConnection.Close()
            '20130129 - set the global count to determine if a message should be generated
            gblFT1Count = FT1Counter
        Catch ex As Exception
            'SendMessage("FT1 Processing Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            'SendMessage("FT1 Processing Complete" & vbCrLf)
        End Try
    End Sub
    Public Sub GeneratePV1(ByRef panum As String)
        Dim sql1 As String = ""
        Dim DFTConnection As New SqlConnection(connString)
        Dim DFTCommand As New SqlCommand
        DFTCommand.Connection = DFTConnection
        Dim dr_DFT As SqlDataReader
        Dim PatientStatusCode As String = ""
        Try
            '20121120 - get the patient status I or O ===============================================================
            sql1 = "select panum, status from [001episode] where panum = '" & panum & "'"
            DFTCommand.CommandText = sql1
            DFTConnection.Open()
            dr_DFT = DFTCommand.ExecuteReader()

            If dr_DFT.HasRows Then

                Do While dr_DFT.Read()
                    If Left(dr_DFT.Item("status"), 1) <> "" Then
                        PatientStatusCode = Left(dr_DFT.Item("Status"), 1)
                    End If

                    strHL7Output = strHL7Output & "PV1||" & PatientStatusCode & "|||||||||||||||||"
                    strHL7Output = strHL7Output & Trim(dr_DFT.Item("panum")) & vbCr

                    '20130228 - Added admindate pv1_44 amnd DCdate pv1_45
                    'strHL7Output = strHL7Output & Trim(dr_DFT.Item("panum").ToString) & "|||||||||||||||||||||||||"
                    'strHL7Output = strHL7Output & buildDate(dr_DFT.Item("adminDate").ToString) & "|" & buildDate(dr_DFT.Item("DCDate").ToString) & vbCrLf

                Loop

            End If
            '20121120 - get the patient status I or O - end ===========================================================
            dr_DFT.Close()
            DFTConnection.Close()

        Catch ex As Exception
            'SendMessage("PV1 Processing Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            'SendMessage("PV1 Processing Complete" & vbCrLf)
        End Try
    End Sub

    Public Sub GeneratePID(ByVal panum As String)
        Dim sql1 As String = ""
        Dim DFTConnection As New SqlConnection(connString)
        Dim DFTCommand As New SqlCommand
        DFTCommand.Connection = DFTConnection
        Dim dr_DFT As SqlDataReader
        Dim intakeFac As String = ""
        Try

            sql1 = "select * from [001episode] where panum = '" & panum & "'"
            DFTCommand.CommandText = sql1
            DFTConnection.Open()
            dr_DFT = DFTCommand.ExecuteReader()

            If dr_DFT.HasRows Then

                Do While dr_DFT.Read()
                    '20121230 changed to use elseif to uniquely identify the intake facility
                    'If dr_DFT("intakeFacility") = 300 Then
                    'intakeFac = "M80"
                    'ElseIf dr_DFT("intakeFacility") = 200 Then
                    'intakeFac = "H80"
                    'Else
                    'intakeFac = ""
                    'End If

                    intakeFac = dr_DFT.Item("star_region")

                    strHL7Output = strHL7Output & "PID|||" & padleft(dr_DFT.Item("mrnum"), 9)
                    'strHL7Output = strHL7Output & "^^^H80|" & padleft(dr_DFT.Item("panum"), 12)
                    'strHL7Output = strHL7Output & "|" & Trim(dr_DFT.Item("lname")) & "^" & Trim(dr_DFT.Item("fname")) & "^" & Trim(dr_DFT.Item("mname")) & "|"
                    strHL7Output = strHL7Output & "^^^" & intakeFac & "^MR||"
                    strHL7Output = strHL7Output & dr_DFT.Item("lname") & "^" & dr_DFT.Item("fname") & "^" & dr_DFT.Item("mname") & "|"


                    strHL7Output = strHL7Output & "|" & buildDate(dr_DFT.Item("DOB")) & "|"
                    strHL7Output = strHL7Output & dr_DFT.Item("gender") & "||" & dr_DFT.Item("race") & "||||||||"
                    strHL7Output = strHL7Output & padleft(dr_DFT.Item("panum"), 9) & "^^^" & intakeFac & "|"
                    strHL7Output = strHL7Output & dr_DFT.Item("socsec") & vbCr
                Loop

            End If

            dr_DFT.Close()
            DFTConnection.Close()
        Catch ex As Exception
            'SendMessage("PID Processing Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            'SendMessage("PID Processing Complete" & vbCrLf)
        End Try
    End Sub
    Public Sub GenerateMSH(ByVal panum As String)
        '20121217 - added code to use panum to ind the intake facility for the MSH segment.
        'process to generate MSH segment for output
        Dim sql1 As String = ""
        Dim DFTConnection As New SqlConnection(connString)
        Dim DFTCommand As New SqlCommand
        DFTCommand.Connection = DFTConnection

        Dim dr_DFT As SqlDataReader
        Dim strYear As String
        Dim strMonth As String
        Dim strDay As String
        Dim strHour As String
        Dim strMinute As String
        Dim strFullDate As String
        Dim theDateTime As String = DateTime.Now
        '05/05/2005 put in code to handle single digit hours and minutes
        '05052005: code to add double bar before ORM^ field

        Try
            sql1 = "select top 1 * from [001episode] where panum = '" & panum & "'"
            DFTCommand.CommandText = sql1
            DFTConnection.Open()
            dr_DFT = DFTCommand.ExecuteReader()

            If dr_DFT.HasRows Then

                Do While dr_DFT.Read()

                    strYear = Year(theDateTime)

                    strMonth = Month(theDateTime)
                    If Len(strMonth) = 1 Then
                        strMonth = "0" & strMonth
                    End If

                    strHour = Hour(theDateTime)
                    If Len(strHour) = 1 Then
                        strHour = "0" & strHour
                    End If

                    strMinute = Minute(theDateTime)
                    If Len(strMinute) = 1 Then
                        strMinute = "0" & strMinute
                    End If

                    strDay = Day(theDateTime)
                    If Len(strDay) = 1 Then
                        strDay = "0" & strDay
                    End If

                    strFullDate = strYear & strMonth & strDay & strHour & strMinute
                    gblStrPostDate = strYear & strMonth & strDay

                    strHL7Output = ""
                    'If dr_DFT.Item("IntakeFacility") = 200 Then
                    'strHL7Output = strHL7Output & "MSH|^~\}|OPS|H80|ITWDFT|H80|" & strFullDate
                    'ElseIf dr_DFT.Item("IntakeFacility") = 300 Then
                    'strHL7Output = strHL7Output & "MSH|^~\}|OPS|M80|ITWDFT|M80|" & strFullDate
                    ' Else

                    'End If

                    '20140405
                    strHL7Output = strHL7Output & "MSH|^~\&|ITW-mCare|" & dr_DFT.Item("star_region") & "|ULH_STAR|" & dr_DFT.Item("star_region") & "|" & strFullDate

                    '05052005: code to add double bar before ORM^ field
                    strHL7Output = strHL7Output & "||DFT^P03|CHPFFINTRANS|P|2.2" & vbCr
                Loop

            End If

            dr_DFT.Close()
            DFTConnection.Close()

        Catch ex As Exception
            'SendMessage("MSH Processing Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            'SendMessage("MSH Processing Complete" & vbCrLf)
        End Try
    End Sub
    Public Sub CreateOutputFile(ByVal strHL7Output As String)
        'Function to create an HL7 output file
        Dim filename As String = ""
        Dim backupFile As String = ""
        Try
            Dim line As String = ""
            Dim objTStreamCounter As Object
            Dim intCounter As Integer = 0
            Dim tempStr As String = Left(strHL7Output, Len(strHL7Output))

            Dim objTStreamOutput As Object

            'If the file does not exist, create it.
            If Not File.Exists(strOutputDirectory & "counter.txt") Then
                objTStreamCounter = File.CreateText(strOutputDirectory & "counter.txt")
                objTStreamCounter.WriteLine("000")
                objTStreamCounter.Close()
            End If

            'read the present file number for counter.Txt. convert it to an integer and increment it.
            objTStreamCounter = New StreamReader(strOutputDirectory & "counter.txt")

            line = objTStreamCounter.readline
            intCounter = CInt(line)
            intCounter = intCounter + 1
            If intCounter >= 100000 Then intCounter = 0
            objTStreamCounter.Close()

            filename = strOutputDirectory & "\HL7." & padleft(Str(intCounter), 3)
            backupFile = strOutputDirectory & "\backup\HL7." & padleft(Str(intCounter), 3)

            objTStreamOutput = File.AppendText(filename)
            '20121119 - add start control character
            objTStreamOutput.Write(Chr(11))

            objTStreamOutput.Write(tempStr)

            '20121119 - add ending control characters
            objTStreamOutput.Write(Chr(28))
            objTStreamOutput.Write(Chr(13))

            objTStreamOutput.Close()

            'update the counter file
            objTStreamCounter = New StreamWriter(strOutputDirectory & "counter.txt")
            objTStreamCounter.WriteLine(padleft(Str(intCounter), 3))
            objTStreamCounter.Close()

            '20130224 - create a backup of the file overwrites existing file

            File.Copy(filename, backupFile, True)

        Catch ex As Exception
            'SendMessage("Create Output File Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            SendMessage(filename & vbCrLf)
        End Try

    End Sub

    Public Function padleft(ByRef inputStr As String, ByRef strLength As Short) As String
        'pad an input string with zeros based on desired strLength
        Dim varLength As Short
        Dim strOutput As String
        Dim i As Short


        strOutput = ""
        varLength = Len(Trim(inputStr))
        For i = 1 To ((strLength - varLength))
            strOutput = strOutput & "0"
        Next
        strOutput = strOutput & Trim(inputStr)
        padleft = strOutput


    End Function

    Function buildDate(ByVal inputStr As String) As String
        Dim strYear As String
        Dim strMonth As String
        Dim strDay As String
        Dim strHour As String
        Dim strMinute As String

        If IsDate(inputStr) Then
            strYear = Trim(Str(Year(inputStr)))
            strMonth = Trim(Str(Month(inputStr)))
            strDay = Trim(Str(Day(inputStr)))
            strHour = (Trim(Str(Hour(inputStr))))
            strMinute = (Trim(Str(Minute(inputStr))))
            If Len(strMonth) = 1 Then
                strMonth = "0" & strMonth
            End If

            If Len(strDay) = 1 Then
                strDay = "0" & strDay
            End If

            If Len(strHour) = 1 Then
                strHour = "0" & strHour
            End If

            If Len(strMinute) = 1 Then
                strMinute = "0" & strMinute
            End If

            If strHour = "00" And strMinute = "00" Then
                buildDate = strYear & strMonth & strDay

            Else
                buildDate = strYear & strMonth & strDay & strHour & strMinute
            End If
        Else
            buildDate = ""
        End If

    End Function

    Public Sub compressfiles(ByVal strLocationPath As String) 'strOutputDirectory & "\backup"

        Dim tempDate As Date = DateTime.Now
        Dim tempYear As String = ""
        Dim tempMonth As String = ""
        Dim tempDay As String = ""
        Dim tempHour As String = ""
        Dim tempMinute As String = ""

        tempMonth = Month(tempDate)
        If Len(tempMonth) = 1 Then
            tempMonth = "0" & tempMonth
        End If

        tempDay = Day(tempDate)
        If Len(tempDay) = 1 Then
            tempDay = "0" & tempDay
        End If

        tempHour = Hour(tempDate)
        If Len(tempHour) = 1 Then
            tempHour = "0" & tempHour
        End If

        tempMinute = Minute(tempDate)
        If Len(tempMinute) = 1 Then
            tempMinute = "0" & tempMinute
        End If


        Dim strZipFileName As String = Year(tempDate) & tempMonth & tempDay & tempHour & tempMinute
        Dim strZipFile = strLocationPath & "\" & strZipFileName & ".zip"

        Dim dirs As String() = Directory.GetFiles(strLocationPath, "HL7.*")

        Using zip As ZipFile = New ZipFile()

            For Each filename In dirs
                Dim theFile As New FileInfo(filename)
                zip.AddFile(filename)
            Next

            zip.Save(strZipFile)
        End Using


        For Each filename In dirs
            Dim theFile As New FileInfo(filename)
            theFile.Delete()
        Next

    End Sub
    Public Sub handleBillingAgent()
        '20130228 - set hl7generated to 1900-01-01 for all records where hl7generated is null and the billing agent for the service code is 1.
        'these are the records that were not processed by the preceeding DFT generation process.
        Dim sqlUpdate As String = ""
        Dim DFTConnection As New SqlConnection(connString)
        Dim DFTCommand As New SqlCommand
        DFTCommand.Connection = DFTConnection
        
        Dim updateConnection As New SqlConnection(connString)
        Dim UpdateCommand As New SqlCommand
        UpdateCommand.Connection = updateConnection

        Dim FT1Counter As Integer = 0
        Dim ft1Segment As String = ""
        Dim intTransactionType As Integer = 0
        Try

            sqlUpdate = "update [075chargeDT] set hl7generated = '1900-01-01' "
            sqlUpdate = sqlUpdate & "where recordID in "
            sqlUpdate = sqlUpdate & "(SELECT [075ChargeDt].recordID "
            sqlUpdate = sqlUpdate & "FROM [075ChargeDt] LEFT OUTER JOIN "
            sqlUpdate = sqlUpdate & "[130Service] ON [075ChargeDt].serviceCode = [130Service].servCode "
            sqlUpdate = sqlUpdate & "WHERE ([075ChargeDt].hl7generated IS NULL)	and ([130Service].billingAgent = 1))"

            UpdateCommand.CommandText = sqlUpdate
            updateConnection.Open()
            UpdateCommand.ExecuteNonQuery()
            updateConnection.Close()

        Catch ex As Exception
            'SendMessage("Billing Agent Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            'SendMessage("FT1 Processing Complete" & vbCrLf)
        End Try
    End Sub
    Public Sub handleFilters()
        '20130417 - handle test cases and deleted episodes. Set date in HL7generated so charges will not be processed later.
        Dim sqlUpdate As String = ""
        Dim DFTConnection As New SqlConnection(connString)
        Dim DFTCommand As New SqlCommand
        DFTCommand.Connection = DFTConnection

        Dim updateConnection As New SqlConnection(connString)
        Dim UpdateCommand As New SqlCommand
        UpdateCommand.Connection = updateConnection

        Dim FT1Counter As Integer = 0
        Dim ft1Segment As String = ""
        Dim intTransactionType As Integer = 0
        Try

            sqlUpdate = "update [075chargeDT] set hl7generated = '1901-01-01' "
            sqlUpdate = sqlUpdate & "where recordID in "
            sqlUpdate = sqlUpdate & "(SELECT c.recordID "
            sqlUpdate = sqlUpdate & "FROM [075ChargeDt] c, [001episode] e "
            sqlUpdate = sqlUpdate & "where c.panum = e.panum "
            sqlUpdate = sqlUpdate & "and e.Status in ('OX','IX') "
            sqlUpdate = sqlUpdate & "and  e.Testcase <> 0)"


            UpdateCommand.CommandText = sqlUpdate
            updateConnection.Open()
            UpdateCommand.ExecuteNonQuery()
            updateConnection.Close()

        Catch ex As Exception
            'SendMessage("Handle Filters Error: " & ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
            Exit Sub
        Finally
            'SendMessage("FT1 Processing Complete" & vbCrLf)
        End Try
    End Sub

    Public Sub CreateErrorFile(ByVal errorstring As String)
        Dim errorfilepath As String = objIniFile.GetString("DFTULH", "ErrorDirULH", "(none)")
        Dim errorfilename As String = String.Format("_DFTError-'" & "'_{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
        Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
        errorfile.Write(errorstring)
        errorfile.Close()
    End Sub

End Module
