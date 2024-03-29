﻿Imports System.Configuration
Imports System.Configuration.ConfigurationSettings

Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.Specialized

Imports System.Data.SqlClient

Public Class DatabaseAccess
    Private sqlConnection As SqlConnection
    Private sqlCommand As SqlCommand

    Public Sub New()
        Dim ConnectionStringToUse As String = ConnectionString()
        Console.WriteLine(ConnectionStringToUse)

        sqlConnection = New SqlConnection(ConnectionStringToUse)
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
            SELECT EMNER.[kunr]
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
              FROM [testAntik].[dbo].[eBayEmner] EMNER, [testAntik].[dbo].[eBayUpdate] UPDATE_TABLE
             WHERE EMNER.[binr] = UPDATE_TABLE.[binr]
               AND UPDATE_TABLE.[eBayUpd] = 1
               AND UPDATE_TABLE.[eBayRettetDD] IS NULL"

        Return sqlCommand.ExecuteReader()
    End Function

    Public Sub UpdateListingResults(binr As Integer, listingFee As Double, eBayProductId As String)
        Dim commandText As String = "
            UPDATE [dbo].[eBayUpdate]
               SET [eBayProductId] = '" & eBayProductId & "'
                  ,[eBayFee] = " & listingFee & "
             WHERE [binr] = " & binr

        ExecuteNonQuery(commandText)
    End Sub

    Public Function FetchCategoryIds(lastCategoryId As Integer) As SqlDataReader
        sqlCommand = sqlConnection.CreateCommand
        sqlCommand.CommandText = "
            SELECT TOP (100) [CategoryID]
              FROM [testAntik].[dbo].[eBayKat]
             WHERE [CategoryID] > " & lastCategoryId

        Return sqlCommand.ExecuteReader()
    End Function

    Public Function FetchItemSpecifics(binr As Integer) As SqlDataReader
        sqlCommand = sqlConnection.CreateCommand
        sqlCommand.CommandText = "
            SELECT [text]
                  ,[customField]
              FROM [testAntik].[dbo].[eBayItemSpecifics]
             WHERE [binr] = " & binr & "
               AND type = 'specific'"

        Return sqlCommand.ExecuteReader()
    End Function

    Private Function ConnectionString() As String
        Return "Server=" & DatabaseServer() & "," & Port() & ";Database=testAntik;Uid=" & UserID() & ";Pwd=" & Password() & ";"
    End Function

    Private Function DatabaseServer() As String
        Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
        Return appSettings.Get("ip_address")
    End Function

    Private Function Port() As String
        Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
        Return appSettings.Get("port")
    End Function

    Private Function UserID() As String
        Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
        Return appSettings.Get("uid")
    End Function

    Private Function Password() As String
        Dim appSettings As NameValueCollection = System.Configuration.ConfigurationManager.AppSettings
        Return appSettings.Get("password")
    End Function
End Class
