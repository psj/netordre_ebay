Imports System
Imports System.Data.SqlClient

Module eBayIntegration

    Sub Main()
        Dim ebayApi As New ebayAPI
        'ebayApi.GetCategoriesAndUpdateDatabase()
        'ebayApi.GetUserInformation()
        'ebayApi.GetCategorySpecifics()

        Dim databaseAccess As New DatabaseAccess
        Dim sqlDataReader As SqlDataReader
        Dim results As String = String.Empty

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
                ItemLocationCountry(sqlDataReader)
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

End Module
