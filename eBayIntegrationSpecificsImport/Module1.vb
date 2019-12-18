Module Module1

    Sub Main()
        Dim databaseAccess As New eBayApiLibrary.eBayApiLibrary.DatabaseAccess
        databaseAccess.OutputConnectionString()

        If databaseAccess.UpdateCategorySpecifics() Then
            Dim ebayApi As New eBayApiLibrary.eBayApiLibrary.ebayAPI
            ebayApi.GetCategorySpecifics()

            databaseAccess.UpdateImportStatus("IS")
        End If
    End Sub

End Module
