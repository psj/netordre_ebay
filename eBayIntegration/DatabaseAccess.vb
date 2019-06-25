Imports System.Data.SqlClient

Public Class DatabaseAccess
    Private sqlConnection As SqlConnection
    Private sqlCommand As SqlCommand

    Public Sub New()
        'sqlConnection = New SqlConnection("Server=localhost\SQLEXPRESS;Database=master;Trusted_Connection=True;")
        sqlConnection = New SqlConnection("Server=195.211.176.180,1433;Database=testAntik; Uid=psj; Pwd=SommerApi2019;")
    End Sub

    Public Function FetchRowsToExport() As SqlDataReader
        sqlCommand = sqlConnection.CreateCommand
        sqlCommand.CommandText = "
            SELECT TOP (10) [kunr]
                  ,[binr]
                  ,[eBayCategoryId]
                  ,[antalEkstraFoto]
                  ,[pris]
                  ,[kort44]
                  ,[lang44]
                  ,[interntnr]
              FROM [testAntik].[dbo].[eBayEmner]"
        sqlConnection.Open()

        Return sqlCommand.ExecuteReader()
    End Function
End Class
