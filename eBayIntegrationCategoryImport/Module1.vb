Module Module1

    Sub Main()
        Dim databaseAccess As New eBayApiLibrary.eBayApiLibrary.DatabaseAccess
        databaseAccess.OutputConnectionString()

        If databaseAccess.UpdateCategories() Then
            Dim ebayApi As New eBayApiLibrary.eBayApiLibrary.ebayAPI
            ebayApi.GetCategoriesAndUpdateDatabase()

            databaseAccess.UpdateImportStatus("Kat")
        End If
    End Sub

End Module
