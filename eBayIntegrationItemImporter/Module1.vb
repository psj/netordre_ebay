Imports System.Data.SqlClient

Module Module1
    ReadOnly ebayApi As New eBayApiLibrary.eBayApiLibrary.ebayAPI

    Sub Main()
        'ebayApi.GetUserInformation()

        Dim databaseAccess As New eBayApiLibrary.eBayApiLibrary.DatabaseAccess
        databaseAccess.OutputConnectionString()

        Dim sqlDataReader As SqlDataReader

        sqlDataReader = databaseAccess.FetchRowsToExport()
        Do While sqlDataReader.Read()
            ebayApi.ListItem(
                BiNr(sqlDataReader),
                CategoryId(sqlDataReader),
                SubCategoryId(sqlDataReader),
                Title(sqlDataReader),
                Description(sqlDataReader),
                Price(sqlDataReader),
                Zoom(sqlDataReader),
                ItemLocationCity(sqlDataReader),
                ItemLocationCountry(sqlDataReader),
                ebayAuthToken(sqlDataReader)
            )
        Loop
        sqlDataReader.Close()

        sqlDataReader = databaseAccess.FetchRowsToUpdate()
        Do While sqlDataReader.Read()
            ebayApi.UpdateItem(
                eBayProductId(sqlDataReader),
                BiNr(sqlDataReader),
                CategoryId(sqlDataReader),
                SubCategoryId(sqlDataReader),
                Title(sqlDataReader),
                Description(sqlDataReader),
                Price(sqlDataReader),
                Zoom(sqlDataReader),
                ItemLocationCity(sqlDataReader),
                ItemLocationCountry(sqlDataReader),
                ebayAuthToken(sqlDataReader)
            )
        Loop
        sqlDataReader.Close()

        sqlDataReader = databaseAccess.FetchRowsToDelete()
        Do While sqlDataReader.Read()
            ebayApi.EndItem(
                eBayProductId(sqlDataReader),
                ebayAuthToken(sqlDataReader)
            )
        Loop
    End Sub

    Function BiNr(row As SqlDataReader) As Integer
        Return row.GetInt32(1)
    End Function

    Function CategoryId(row As SqlDataReader) As Integer
        Return row.GetInt32(2)
    End Function

    Function SubCategoryId(row As SqlDataReader) As Integer
        Try
            Return row.GetInt32(3)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function Title(row As SqlDataReader) As String
        Return row.GetString(6)
    End Function

    Function Description(row As SqlDataReader) As String
        Return row.GetString(7)
    End Function

    Function Price(row As SqlDataReader) As Integer
        Try
            Return row.GetInt32(5)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function Zoom(row As SqlDataReader) As Boolean
        Return row.GetBoolean(9)
    End Function

    Function ItemLocationCity(row As SqlDataReader) As String
        Return row.GetString(12)
    End Function

    Function ItemLocationCountry(row As SqlDataReader) As String
        Return row.GetString(13)
    End Function

    Function eBayProductId(row As SqlDataReader) As String
        Return row.GetString(14)
    End Function

    Function ebayAuthToken(row As SqlDataReader) As String
        If ebayApi.IsProduction() Then
            Return row.GetString(15)
        Else
            Return row.GetString(16)
        End If
    End Function

    Function CustomerAboutText(row As SqlDataReader) As String
        Return row.GetString(18)
    End Function


End Module
