Imports System.Data.SqlClient

Public Class DatabaseAccess
    Private sqlConnection As SqlConnection
    Private sqlCommand As SqlCommand

    Public Sub New()
        sqlConnection = New SqlConnection("Server=localhost\SQLEXPRESS;Database=master;Trusted_Connection=True;")
    End Sub

    Public Function FetchRowsToExport() As SqlDataReader
        sqlCommand = sqlConnection.CreateCommand
        sqlCommand.CommandText = "SELECT tekst FROM [ebay_integration].[dbo].[test]"
        sqlConnection.Open()

        Return sqlCommand.ExecuteReader()
    End Function
End Class
