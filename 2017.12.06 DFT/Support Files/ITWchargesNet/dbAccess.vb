Imports System.Data
Imports System.Data.SqlClient

Public Class dbAccess
    Private Shared Connection_String As String = "data source=" & System.Configuration.ConfigurationSettings.AppSettings("dbLocation") & ";initial catalog=" & System.Configuration.ConfigurationSettings.AppSettings("dbCategory") & ";user id=sysmax;password=Condor!;persist security info=True;workstation id=mattdev;packet size=4096"
    Private Shared Connection As SqlConnection = New SqlConnection(Connection_String)

    Public Shared Sub TestConnection()
        Try
            Connection.Open()
        Catch ex As Exception
            Throw ex
        Finally
            Connection.Close()
        End Try
    End Sub

    ' Function to get a DataReader based on an sql string
    Public Shared Function getDataReader(ByVal sql As String) As SqlDataReader
        Dim dataReader As SqlDataReader = Nothing
        Dim comm As SqlCommand = Nothing

        Try
            ' Open the connection
            Connection.Open()

            ' Create the Command object
            comm = New SqlCommand(sql, Connection)
            comm.CommandTimeout = 60
            ' Create the DataReader object
            dataReader = comm.ExecuteReader(CommandBehavior.CloseConnection)

            ' Cannot close the connection because the data reader relies upon it
            ' By using CommandBehavior.CloseConnection above, the connection will
            '  be closed when the dataReader is closed
            'Connection.Close()
        Catch ex As Exception
            ' Display and log any errors
            Throw ex
        Finally
            comm = Nothing
        End Try

        ' Return the DataReader
        Return dataReader
    End Function

    Public Shared Function getDataSet(ByVal sql As String) As System.Data.DataSet
        Return getDataSet(sql, "theTable")
    End Function

    ' Function to get a DataSet based on an sql string
    Public Shared Function getDataSet(ByVal sql As String, ByVal theName As String) As System.Data.DataSet
        Dim dataSet As System.Data.DataSet = Nothing
        Dim selectCommand As SqlCommand = Nothing
        Dim dataAdapter As SqlDataAdapter = Nothing

        Try
            ' Set up the Command, DataAdapter, and DataSet objects necessary
            selectCommand = New SqlCommand(sql, Connection)
            selectCommand.CommandTimeout = 60
            dataAdapter = New SqlDataAdapter(selectCommand)
            dataSet = New System.Data.DataSet

            ' Use the DataAdapter to fill the DataSet with the data
            dataAdapter.Fill(dataSet, theName)
        Catch ex As Exception
            ' Display and log any errors
            Throw ex
        Finally
            dataAdapter = Nothing
            selectCommand = Nothing
        End Try

        ' Return the DataSet
        Return dataSet
    End Function

    Public Shared Function fillDataSetTable(ByRef ds As DataSet, ByVal sql As String, ByVal tableName As String) As Boolean
        Dim selectCommand As SqlCommand = Nothing
        Dim dataAdapter As SqlDataAdapter = Nothing

        Try
            ' Set up the Command and DataAdapter
            selectCommand = New SqlCommand(sql, Connection)
            ' 06/14/2007 Per Max, allow the query to go for up to 2 minutes
            ' 10/10/2007 Make that 10 minutes since the exclusion query was
            '  taking about 6 or 7 minutes during the day
            selectCommand.CommandTimeout = 600 '120
            dataAdapter = New SqlDataAdapter(selectCommand)

            dataAdapter.Fill(ds, tableName)
        Catch ex As Exception
            ' Display and log any errors
            Throw ex
        Finally
            dataAdapter = Nothing
            selectCommand = Nothing
        End Try

        Return True
    End Function

    ' Funtion to run a query from an sql string
    Public Shared Function executeNonQuery(ByVal sql As String) As Boolean
        Dim updateCommand As SqlCommand = Nothing

        Try
            ' Create an OleDb Command used to run the query
            updateCommand = New SqlCommand(sql, Connection)
            updateCommand.CommandTimeout = 60

            ' Open the connection
            Connection.Open()
            ' Use the ExecuteNonQuery() method to run the query
            updateCommand.ExecuteNonQuery()

            ' Return True for success
            Return True
        Catch ex As Exception
            ' Display and log any errors
            ' Return False for failure
            Throw ex
        Finally
            ' Close the connection
            Connection.Close()
            updateCommand = Nothing
        End Try
    End Function

    ' Function to get a DataAdapter from an sql string
    Public Shared Function getDataAdapter(ByVal sql As String) As SqlDataAdapter
        Try
            ' Return the DataAdapter
            Return New SqlDataAdapter(sql, Connection)
        Catch ex As Exception
            ' Display and log any errors
            Throw ex
        End Try
    End Function
End Class
