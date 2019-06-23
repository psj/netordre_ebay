Imports System
Imports System.Data.SqlClient

Module eBayIntegration

    Sub Main()
        Dim ebayAuthentication As EbayAuthentication = New EbayAuthentication

        Dim databaseAccess As DatabaseAccess
        Dim sqlDataReader As SqlDataReader
        Dim results As String

        results = String.Empty
        databaseAccess = New DatabaseAccess
        sqlDataReader = databaseAccess.FetchRowsToExport()

        Do While sqlDataReader.Read()
            results = results & sqlDataReader.GetString(0) & vbTab
        Loop

    End Sub

End Module
