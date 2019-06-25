Imports System
Imports System.Data.SqlClient

Module eBayIntegration

    Sub Main()
        Logger.LogInfo("eBayIntegration starting up")

        Dim ebayApi As New ebayAPI
        ebayApi.GetUserInformation()
        ebayApi.ListItem()

        Dim databaseAccess As New DatabaseAccess
        Dim sqlDataReader As SqlDataReader
        Dim results As String = String.Empty

        sqlDataReader = databaseAccess.FetchRowsToExport()
        Do While sqlDataReader.Read()
            results = results & sqlDataReader.GetInt32(0).ToString() & ";" & sqlDataReader.GetInt32(1) & ";"
        Loop
    End Sub

End Module
