Module Module1

    Sub Main()
        Dim ebayApi As New eBayApiLibrary.eBayApiLibrary.ebayAPI
        ebayApi.GetCategoriesAndUpdateDatabase()
    End Sub

End Module
