Imports System.Configuration
Imports System.Configuration.ConfigurationSettings

Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.Specialized

Imports System.Data.SqlClient

Namespace eBayApiLibrary
    Public Class DatabaseAccess
        Private sqlConnection As SqlConnection
        Private sqlCommand As SqlCommand

        Public Sub New()
            sqlConnection = New SqlConnection(ConnectionString())
            sqlConnection.Open()
        End Sub

        Public Sub Dispose()
            sqlConnection.Close()
            sqlConnection.Dispose()
        End Sub

        Public Sub InsertOrUpdateCategory(BestOfferEnabled As Boolean, AutoPayEnabled As Boolean, CategoryID As Integer, CategoryLevel As Integer, CategoryName As String, CategoryParentID As Integer)
            Dim sanitizedCategoryName = CategoryName.Replace("'", "''")
            Dim commandText As String = "INSERT INTO [testAntik].[dbo].[eBayKat] ([BestOfferEnabled], [AutoPayEnabled], [CategoryID], [CategoryLevel], [CategoryName], [CategoryParentID]) VALUES (" & ConvertBooleanToIntString(BestOfferEnabled) & "," & ConvertBooleanToIntString(AutoPayEnabled) & "," & CategoryID.ToString & "," & CategoryLevel.ToString & ",'" & sanitizedCategoryName & "'," & CategoryParentID.ToString & ")"
            ExecuteNonQuery(commandText)
        End Sub

        Public Sub DeleteExistingCategories()
            Dim commandText As String = "DELETE From [testAntik].[dbo].[eBayKat]"
            ExecuteNonQuery(commandText)
        End Sub

        Public Sub MergeCategoriesFromTest()
            Dim commandText As String = MergeCategoriesCommandText()
            ExecuteNonQuery(commandText)
        End Sub

        Public Sub MergeSpecificsFromTest()
            Dim commandText As String = MergeSpecificsCommandText()
            ExecuteNonQuery(commandText)
        End Sub

        Public Sub InsertCategorySpecific(CategoryId As Integer, Name As String, Mandatory As Boolean, ValueType As String, ValueRecommendations As String)
            Dim sanitizedName = Name.Replace("'", "''")
            Dim sanitizedValueRecommendations = ValueRecommendations.Replace("'", "''")
            Dim commandText As String = "INSERT INTO [testAntik].[dbo].[eBayItemSpecificDefinitions] ([CategoryId], [name], [mandatory], [valueType], [valueRecommendations]) VALUES (" & CategoryId.ToString & ",'" & sanitizedName & "'," & ConvertBooleanToIntString(Mandatory) & "," & ValueType & ",'" & sanitizedValueRecommendations & "')"
            ExecuteNonQuery(commandText)
        End Sub

        Public Sub DeleteExistingCategorySpecifics()
            Dim commandText As String = "DELETE From [testAntik].[dbo].[eBayItemSpecificDefinitions]"
            ExecuteNonQuery(commandText)
        End Sub

        Private Sub ExecuteNonQuery(CommandText As String)
            Dim nonQuerySqlCommand As SqlCommand
            nonQuerySqlCommand = sqlConnection.CreateCommand
            nonQuerySqlCommand.CommandText = CommandText
            nonQuerySqlCommand.ExecuteNonQuery()
        End Sub

        Private Function ConvertBooleanToIntString(booleanValue As Boolean) As String
            If booleanValue = True Then
                Return "1"
            Else
                Return "0"
            End If
        End Function

        Public Function FetchRowsToExport() As SqlDataReader
            sqlCommand = sqlConnection.CreateCommand
            sqlCommand.CommandText = "
            SELECT " & ColumnsToFetch() & "
              FROM [eBayEmner] EMNER
             WHERE EMNER.[eBayUpd] = 1
               AND EMNER.[eBayProductId] IS NULL"

            Return sqlCommand.ExecuteReader()
        End Function

        Public Function FetchRowsToUpdate() As SqlDataReader
            sqlCommand = sqlConnection.CreateCommand
            sqlCommand.CommandText = "
            SELECT " & ColumnsToFetch() & "
              FROM [eBayEmner] EMNER
             WHERE EMNER.[eBayUpd] = 1
               AND EMNER.[eBayProductId] IS NOT NULL"

            Return sqlCommand.ExecuteReader()
        End Function

        Public Function FetchRowsToDelete() As SqlDataReader
            sqlCommand = sqlConnection.CreateCommand
            sqlCommand.CommandText = "
            SELECT " & ColumnsToFetch() & "
              FROM [eBayEmner] EMNER
             WHERE EMNER.[eBaySlet] = 1
               AND EMNER.[eBayProductId] IS NOT NULL"

            Return sqlCommand.ExecuteReader()
        End Function

        Private Function ColumnsToFetch() As String
            Return "EMNER.[kunr]
                   ,EMNER.[binr]
                   ,EMNER.[eBayCategoryId]
                   ,EMNER.[eBayCategoryId2]
                   ,EMNER.[antalEkstraFoto]
                   ,EMNER.[pris]
                   ,EMNER.[kort44]
                   ,EMNER.[lang44]
                   ,EMNER.[interntnr]
                   ,EMNER.[apZoom]
                   ,EMNER.[enhed]
                   ,EMNER.[enhed44]
				   ,EMNER.[lokation_city]
				   ,EMNER.[lokation_country]
                   ,EMNER.[eBayProductId]
                   ,EMNER.[Token]
                   ,EMNER.[TokenAppIdSandbox]"
        End Function

        Public Sub UpdateListingResults(binr As Integer, listingFee As Double, eBayProductId As String)
            Dim commandText As String = "
            UPDATE [dbo].[eBayUpdate]
               SET [eBayProductId] = '" & eBayProductId & "'
                  ,[eBayUpd] = 0
                  ,[eBayLog] = NULL
                  ,[eBayError] = 0
                  ,[eBayRettetDD] = GETDATE()
                  ,[eBayFee] = " & listingFee & "
             WHERE [binr] = " & binr

            ExecuteNonQuery(commandText)
        End Sub

        Public Sub UpdateRevisedListingResults(binr As Integer)
            Dim commandText As String = "
            UPDATE [dbo].[eBayUpdate]
               SET [eBayUpd] = 0
                  ,[eBayLog] = NULL
                  ,[eBayError] = 0
                  ,[eBayRettetDD] = GETDATE()
             WHERE [binr] = " & binr

            ExecuteNonQuery(commandText)
        End Sub

        Public Sub UpdateEndListingResults(eBayProductId As String)
            Dim commandText As String = "
            UPDATE [dbo].[eBayUpdate]
               SET [eBaySlet] = 0
                  ,[eBaySletDD] = GETDATE()
                  ,[eBayProductId] = NULL
             WHERE [eBayProductId] = " & eBayProductId

            ExecuteNonQuery(commandText)
        End Sub

        Public Sub UpdateListingError(binr As Integer, ErrorMessage As String)
            Dim commandText As String = "
            UPDATE [dbo].[eBayUpdate]
               SET [eBayError] = 1
                  ,[eBayLog] = '" & ErrorMessage & "'
             WHERE [binr] = " & binr

            ExecuteNonQuery(commandText)
        End Sub

        Public Function FetchCategoryIds(lastCategoryId As Integer) As SqlDataReader
            sqlCommand = sqlConnection.CreateCommand
            sqlCommand.CommandText = "
            SELECT TOP (100) [CategoryID]
              FROM [dbo].[eBayKat]
             WHERE [CategoryID] > " & lastCategoryId

            Return sqlCommand.ExecuteReader()
        End Function

        Public Function FetchItemSpecifics(binr As Integer) As SqlDataReader
            sqlCommand = sqlConnection.CreateCommand
            sqlCommand.CommandText = "
            SELECT [type]
                  ,[text]
                  ,[customField]
              FROM [dbo].[eBayItemSpecifics]
             WHERE [binr] = " & binr

            Return sqlCommand.ExecuteReader()
        End Function

        Public Function UpdateCategories() As Boolean
            sqlCommand = sqlConnection.CreateCommand
            sqlCommand.CommandText = "
            SELECT [updKat]
              FROM [dbo].[eBayGet]
             WHERE [updKatDD] IS NULL"

            Dim sqlDataReader As SqlDataReader
            sqlDataReader = sqlCommand.ExecuteReader()

            If sqlDataReader.HasRows() Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function UpdateCategorySpecifics() As Boolean
            sqlCommand = sqlConnection.CreateCommand
            sqlCommand.CommandText = "
            SELECT [updIS]
              FROM [dbo].[eBayGet]
             WHERE [updISdd] IS NULL"

            Dim sqlDataReader As SqlDataReader
            sqlDataReader = sqlCommand.ExecuteReader()

            If sqlDataReader.HasRows() Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub UpdateImportStatus(fieldName As String)
            Dim commandText As String = "
            UPDATE [dbo].[eBayGet]
               SET [upd" & fieldName & "DD] = GETDATE()
             WHERE [upd" & fieldName & "] = 1"

            ExecuteNonQuery(commandText)
        End Sub

        Public Sub OutputConnectionString()
            Console.WriteLine(PrintConnectionString())
        End Sub

        Private Function ConnectionString() As String
            Return PrintConnectionString() & ";Pwd=" & Password() & ";"
        End Function

        Private Function PrintConnectionString() As String
            Return "Server=" & DatabaseServer() & "," & Port() & ";Database=" & DatabaseName() & ";Uid=" & UserID()
        End Function

        Private Function DatabaseServer() As String
            Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
            Return appSettings.Get("ip_address")
        End Function

        Private Function Port() As String
            Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
            Return appSettings.Get("port")
        End Function

        Private Function DatabaseName() As String
            Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
            Return appSettings.Get("database")
        End Function

        Private Function UserID() As String
            Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
            Return appSettings.Get("uid")
        End Function

        Private Function Password() As String
            Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
            Return appSettings.Get("password")
        End Function

        Private Function MergeSpecificsCommandText() As String
            Return "MERGE phd400.dbo.eBayItemSpecificDefinitions AS T
                    USING (SELECT categoryId, [name], mandatory, valueType, valueRecommendations FROM eBayItemSpecificDefinitions) AS S
                    ON (T.CategoryID = S.CategoryID) AND T.[name] = S.[name]
                    WHEN MATCHED THEN UPDATE SET T.mandatory = S.mandatory, T.valueType = S.valueType, T.valueRecommendations = S.valueRecommendations
                    WHEN NOT MATCHED BY TARGET THEN INSERT (categoryId, [name], mandatory, valueType, valueRecommendations)
                                                 VALUES (S.categoryId, S.[name], S.mandatory, S.valueType, S.valueRecommendations)
                    WHEN NOT MATCHED BY SOURCE THEN DELETE
                    ;"
        End Function

        Private Function MergeCategoriesCommandText() As String
            Return "MERGE phd400.dbo.eBayKat AS T
                    USING (SELECT BestOfferEnabled, AutoPayEnabled, CategoryID, CategoryLevel, CategoryName, CategoryParentID FROM testAntik.dbo.eBayKat) AS S
                    ON (T.CategoryID = S.CategoryID)
                    WHEN MATCHED THEN UPDATE SET T.BestOfferEnabled = S.BestOfferEnabled, T.AutoPayEnabled = S.AutoPayEnabled,T.CategoryLevel = S.CategoryLevel
                                                 , T.CategoryName = S.CategoryName, T.CategoryParentID = S.CategoryParentID
                    WHEN NOT MATCHED BY TARGET THEN INSERT (BestOfferEnabled, AutoPayEnabled, CategoryID, CategoryLevel, CategoryName, CategoryParentID)
                                                 VALUES (S.BestOfferEnabled, S.AutoPayEnabled, S.CategoryID, S.CategoryLevel, S.CategoryName, S.CategoryParentID)
                    WHEN NOT MATCHED BY SOURCE THEN DELETE
                    ;"
        End Function
    End Class
End Namespace