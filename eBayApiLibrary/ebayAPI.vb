Imports eBay.Service.Call
Imports eBay.Service.Core.Sdk
Imports eBay.Service.Core.Soap
Imports eBay.Service.Util
Imports eBay.Service.EPS
Imports System.DateTime
Imports System.Data.SqlClient

Namespace eBayApiLibrary
    Public Class ebayAPI
        Private apiContext As ApiContext = Nothing
        Private ApiServerUrl As String = "https://api.sandbox.ebay.com/wsapi"

        ' testuser ??
        'Private ebayAuthToken As String = "AgAAAA**AQAAAA**aAAAAA**CojZXA**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6wFk4aiCZKFoQydj6x9nY+seQ**7/4EAA**AAMAAA**saEELQopmdyEDeta6ILFQHmzZG9lK8nwKD57KDJHv49DCHEb96nEkqAnpuslD+Wuv/R55NXHE1UXqpj6C/yzglgNaaK13Bd2wNV/Gy//Ydh/ZEuRs1jtgURFw+rEfj/0fFRhXQH+XFPeUhfRmAkQRC7RKw/RcBP7i63y/LVa3KMDQmEhFrHK8+uDk+L+guT948IAHpWck+iN65WozrxU6TmqQ3tk6ZCWQXjNw4hYFJRi19HSI43MuBvRi0wV8hg7evaCu5WCPSNY2cTZbv5BagLIRGcD6h8dCCUotakEfqIun0qexnuAigeaNvLo8Cneoz5WNV7ZYbr/ynZelNDTSpUANDpNrAN2tsirWBThYI3RU/XYJgZUrN8Z7d1wPSpbnKfVfrSvtzkAdnh64mMz1sxUlPUwcVhZHWsa69EBk9sIkqXjbtq5ERxhzgbx2DqwK1foUAf6X7+W5a2/wsF24ERq3Dg1pyPtXe0+pXI6Mp6iNuG9onsLOx2lPw53GsBerPD2hFKh6nxB8m7tTUPagz1ti5Eg8GOpsmHIlq5mNX/ZgOc4nxnimMhyCgErtFCsdEYOeBP9wuUoz2v6hc/7OAc0Ax5j3Ev8yXbkW9cwSKEHzIP4yH22tZue6pKsVIYRk1Z2bHOynIrbkFi5tPztyGyPFkAcUhZertXWG4LtnPkJR/CjRuPsY9ekP3vs6f22nRn03/qQDy/o4GIQWtFzS4+qfUpPLmYIJ1KDck45dLUNu8oZsc5VWkJjSyrQgUQr"

        ' testuser_petersandager
        Private ebayAuthToken As String = "AgAAAA**AQAAAA**aAAAAA**H6glXQ**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6wFk4aiCZWEpAqdj6x9nY+seQ**DAcFAA**AAMAAA**EINwVuiMOj7p00p81bbVXgnbFO4JAPtp0rxdLYfIerKxFyl0IJIVfti1LPT9Ogqgcg/51W7YIFJHCwrQS3sedOr5/Ed5iZBoe8MCLSDPbyvWutPNoTpSRCBx5MJVm0ZYp9PLTLiFNW5gZOReAtIO9RkbHr3g8gB1pMkVj4X+vUOPKZxX+u3uT0SsZTZxw5UUGMnxnrUn0coyzSr7MOZWk3l/150Uu5Z7xuJf/iLVMN+fZoHD2tpA48P719IJqv2qokvapWrViDMZhv6lEKhf4dyExyReKnwdbdRKAB4uTChywFClp2h173u5kEEKMVlDqH0CwempMn1EhOnrNBk0FCdNZXIiPmryWx5DsU4Y/XDf5s7pX68ArCdKWOn0XBtHEX0jjgsRfnFNIgJlX+/f9kBtnKGFv511DDvYPb/A5gKbxdZMIK4Tleg6CskTeRY7gOO7UCjNBSKFbrqjEVw3i1dI4X5dn9Cu1FrtRrpwJcoaFPhXcyjf2WRX3O7Qlr7XBbjx4nXSTDC/h9v89SdWroXiRZzCOm5NwoJio2yCJXj3G2le/dzefnnTQJxpQ6t4A2DP1NSplMiLLEq5bk++URwVjikVWHCJHsqVLvljhhlHTaFBXDpQavOdn+9trCuw668ch6cRpjAAasRWW3Y6xFZK9I/hZh6jDoV9FqlFTPrwjl9Ma/RjSjB44Kw5eay9M4ZbJmH9Lgz2251v/zqDIW1qU+v6zmArxoJ5+CYb/nfmcWQphjBHKg/XQdwGbNEL"

        Public Sub New()
            GetApiContext()
        End Sub

        Public Sub GetCategoriesAndUpdateDatabase()
            Try
                Dim apiCall As GetCategoriesCall = New GetCategoriesCall(apiContext)
                apiCall.DetailLevelList.Add(DetailLevelCodeType.ReturnAll)
                apiCall.Execute()

                Dim databaseAccess As New DatabaseAccess
                Dim categories As CategoryTypeCollection = apiCall.GetCategories()
                Dim category As CategoryType

                For Each category In categories
                    If category.CategoryLevel = 1 Or category.CategoryLevel = 0 Then
                        Console.WriteLine(category.CategoryID.ToString + ", " + category.CategoryLevel.ToString + ", " + category.CategoryName)
                    End If
                    databaseAccess.InsertOrUpdateCategory(category.BestOfferEnabled, category.AutoPayEnabled, category.CategoryID, category.CategoryLevel, category.CategoryName, category.CategoryParentID(0))
                Next category

            Catch ex As Exception
                Console.WriteLine("Failed to get categories : " + ex.Message)
            End Try
        End Sub

        Public Sub ListItem(binr As Integer, CategoryId As Integer, SubCategoryId As Integer, title As String, description As String, price As Integer, zoom As Boolean, ItemLocationCity As String, ItemLocationCountry As String)
            Try
                Dim item As ItemType = BuildItem(binr, CategoryId, SubCategoryId, title, description, price, zoom, ItemLocationCity, ItemLocationCountry)

                Dim apiCall As AddFixedPriceItemCall = New AddFixedPriceItemCall(apiContext)
                Dim fees As FeeTypeCollection = apiCall.AddFixedPriceItem(item)

                Console.WriteLine("The item was listed successfully!")

                Dim listingFee As Double = 0.0
                Dim fee As FeeType
                For Each fee In fees
                    If (fee.Name = "ListingFee") Then
                        listingFee = fee.Fee.Value
                    End If
                Next

                Dim databaseAccess As New DatabaseAccess
                databaseAccess.UpdateListingResults(binr, listingFee, item.ItemID)

                Console.WriteLine(String.Format("Listing fee is: {0}", listingFee))
                Console.WriteLine(String.Format("Listed Item ID: {0}", item.ItemID))
            Catch ex As Exception
                Console.WriteLine("Failed to list item : " + ex.Message)
            End Try

        End Sub

        Public Sub UpdateItem(eBayProductId As String, binr As Integer, CategoryId As Integer, SubCategoryId As Integer, title As String, description As String, price As Integer, zoom As Boolean, ItemLocationCity As String, ItemLocationCountry As String)
            Try
                Dim item As ItemType = BuildItem(binr, CategoryId, SubCategoryId, title, description, price, zoom, ItemLocationCity, ItemLocationCountry)
                item.ItemID = eBayProductId

                Dim apiCall As ReviseFixedPriceItemCall = New ReviseFixedPriceItemCall(apiContext)

                Dim deletedFields As eBay.Service.Core.Soap.StringCollection = New StringCollection()
                Dim fees As FeeTypeCollection = apiCall.ReviseFixedPriceItem(item, deletedFields)

                Console.WriteLine("The item was updated successfully!")

                Dim listingFee As Double = 0.0
                Dim fee As FeeType
                For Each fee In fees
                    If (fee.Name = "ListingFee") Then
                        listingFee = fee.Fee.Value
                    End If
                Next

                Dim databaseAccess As New DatabaseAccess
                databaseAccess.UpdateListingResults(binr, listingFee, item.ItemID)

                Console.WriteLine(String.Format("Listing fee is: {0}", listingFee))
                Console.WriteLine(String.Format("Listed Item ID: {0}", item.ItemID))
            Catch ex As Exception
                Console.WriteLine("Failed to update item : " + ex.Message)
            End Try

        End Sub

        Public Sub EndItem(eBayProductId As String)
            Try
                Dim apicall As EndItemCall = New EndItemCall(apiContext)
                apicall.EndItem(eBayProductId, EndReasonCodeType.NotAvailable)

                Dim databaseAccess As New DatabaseAccess
                databaseAccess.UpdateEndListingResults(eBayProductId)
            Catch ex As Exception
                Console.WriteLine("Failed to end item : " + ex.Message)
            End Try
        End Sub

        Private Function BuildItem(binr As Integer, CategoryId As Integer, SubCategoryId As Integer, title As String, description As String, price As Integer, zoom As Boolean, ItemLocationCity As String, ItemLocationCountry As String) As ItemType
            If price = 0 Then
                Throw New System.Exception("Price cannot be 0")
            End If

            Dim item As ItemType = New ItemType With {
            .Title = title,
            .Description = description,
            .BuyerGuaranteePrice = NewAmount(price),
            .StartPrice = NewAmount(price),
            .ListingType = ListingTypeCodeType.FixedPriceItem,
            .Currency = CurrencyCodeType.USD,
            .ListingDuration = "GTC",
            .Location = ItemLocationCity
        }

            Dim CountryCodeTranslator As CountryCodePolicy = New CountryCodePolicy
            item.Country = CountryCodeTranslator.TranslateCountryCodeToId(ItemLocationCountry)

            Dim category As CategoryType = New CategoryType With {
            .CategoryID = CategoryId.ToString
        }

            item.PrimaryCategory = category

            If Not SubCategoryId = 0 Then
                Dim subCategory As CategoryType = New CategoryType With {
                .CategoryID = SubCategoryId.ToString
            }
                item.SecondaryCategory = subCategory
            End If

            item.PaymentMethods = New BuyerPaymentMethodCodeTypeCollection From {
            BuyerPaymentMethodCodeType.PayPal
        }
            item.PayPalEmailAddress = "me@paypal.com"

            item.DispatchTimeMax = 1
            item.ShippingDetails = BuildShippingDetails()

            item.ReturnPolicy = New ReturnPolicyType With {
            .ReturnsAcceptedOption = "ReturnsAccepted",
            .ReturnsWithin = 30
        }

            item.PictureDetails = New PictureDetailsType With {
            .PhotoDisplaySpecified = True,
            .PhotoDisplay = PhotoDisplayCodeType.None,
            .PictureURL = New StringCollection()
        }

            If zoom = True Then
                item.PictureDetails.PictureURL.Add("https://www.antikvitet.net/images/apzoom/" & binr & "z.jpg")
            Else
                item.PictureDetails.PictureURL.Add("https://www.antikvitet.net/images/antLarge/" & binr & ".jpg")
            End If

            item.ItemSpecifics = BuildItemSpecifics(binr)

            Return item
        End Function

        Private Function BuildItemSpecifics(binr As Integer) As NameValueListTypeCollection
            Dim databaseAccess As New DatabaseAccess
            Dim sqlDataReader As SqlDataReader
            Dim results As String = String.Empty

            Dim nvCollection As NameValueListTypeCollection = New NameValueListTypeCollection()
            sqlDataReader = databaseAccess.FetchItemSpecifics(binr)
            'sqlDataReader = databaseAccess.FetchItemSpecifics(134119)

            Do While sqlDataReader.Read()
                Dim specificsText As String = sqlDataReader.GetString(0)

                Dim nameValueListType As NameValueListType = New NameValueListType()
                nameValueListType.Name = specificsText.Split(",").First

                Dim nvStringCollection As StringCollection = New StringCollection()
                Dim stringArray() As String = {specificsText.Split(",").Last}
                nvStringCollection.AddRange(stringArray)
                nameValueListType.Value = nvStringCollection
                nvCollection.Add(nameValueListType)

                'nameValueListType.Name = "Object Type"

                'nvStringCollection.Clear()
                'stringArray = {"Ashtray"}
                'nvStringCollection.AddRange(stringArray)
                'nameValueListType.Value = nvStringCollection
                'nvCollection.Add(nameValueListType)
            Loop

            Return nvCollection
        End Function

        Private Function BuildShippingDetails() As ShippingDetailsType
            Dim shippingDetails As ShippingDetailsType = New ShippingDetailsType With {
            .ShippingType = ShippingTypeCodeType.Flat
        }

            'shippingOptions.ShippingInsuranceCost = NewAmount(1)
            Dim shippingOptionsWorldwide As ShippingServiceOptionsType = New ShippingServiceOptionsType With {
            .ShippingService = ShippingServiceCodeType.ShippingMethodStandard.ToString(),
            .ShippingServiceCost = NewAmount(30),
            .ShippingServicePriority = 1,
            .LocalPickupSpecified = True,
            .LocalPickup = True
        }

            Dim shippingOptionsLocalPickup As ShippingServiceOptionsType = New ShippingServiceOptionsType With {
            .ShippingService = ShippingServiceCodeType.Pickup.ToString(),
            .ShippingServiceCost = NewAmount(0),
            .LocalPickupSpecified = True,
            .LocalPickup = True
        }

            shippingDetails.ShippingServiceOptions = New ShippingServiceOptionsTypeCollection From {
            shippingOptionsWorldwide,
            shippingOptionsLocalPickup
        }

            Return shippingDetails
        End Function

        Public Sub GetCategorySpecifics()
            Dim mandatorySpecific As Boolean
            Dim lastCategoryId As Integer = 0

            Dim databaseAccess As New DatabaseAccess
            databaseAccess.DeleteExistingCategorySpecifics()

            Dim apiCall As GetCategorySpecificsCall = New GetCategorySpecificsCall(apiContext)

            Dim categoryIdList As StringCollection = New StringCollection
            Dim emptyCollection As CategoryItemSpecificsTypeCollection = New CategoryItemSpecificsTypeCollection
            Dim categorySpecifics As RecommendationsTypeCollection

            Dim categorySpecific As RecommendationsType
            Dim nameRecommendation As NameRecommendationType
            Dim valueRecommendation As ValueRecommendationType
            Dim valueRecommendations As String

            Do While True
                Console.WriteLine(lastCategoryId.ToString())

                categoryIdList = FetchCategoryList(lastCategoryId)
                If categoryIdList.Count() = 0 Then
                    Return
                End If

                categorySpecifics = apiCall.GetCategorySpecifics(categoryIdList)

                For Each categorySpecific In categorySpecifics
                    For Each nameRecommendation In categorySpecific.NameRecommendation
                        'Console.WriteLine(nameRecommendation.Name)
                        If nameRecommendation.ValidationRules.MinValues > 0 Then
                            mandatorySpecific = True
                        Else
                            mandatorySpecific = False
                        End If

                        valueRecommendations = String.Empty
                        For Each valueRecommendation In nameRecommendation.ValueRecommendation
                            If valueRecommendations = String.Empty Then
                                valueRecommendations = valueRecommendation.Value
                            Else
                                valueRecommendations += "," + valueRecommendation.Value
                            End If
                        Next valueRecommendation

                        databaseAccess.InsertCategorySpecific(categorySpecific.CategoryID, nameRecommendation.Name, mandatorySpecific, nameRecommendation.ValidationRules.ValueType, valueRecommendations)

                    Next nameRecommendation

                    lastCategoryId = categorySpecific.CategoryID
                Next categorySpecific
            Loop
        End Sub

        Private Function FetchCategoryList(lastCategoryId As Integer) As StringCollection
            Dim databaseAccess As New DatabaseAccess
            Dim sqlDataReader As SqlDataReader
            sqlDataReader = databaseAccess.FetchCategoryIds(lastCategoryId)

            Dim categoryIdList As StringCollection = New StringCollection
            Dim i As Integer

            i = 0
            Do While i < 100
                If sqlDataReader.Read() = False Then
                    Return categoryIdList
                Else
                    categoryIdList.Add(sqlDataReader.GetInt32(0).ToString)
                End If

                i += 1
            Loop

            sqlDataReader.Close()
            Return categoryIdList
        End Function

        Public Sub GetUserInformation()
            Try
                Dim apiCall As GetUserCall = New GetUserCall(apiContext)
                apiCall.DetailLevelList.Add(DetailLevelCodeType.ReturnAll)

                Console.WriteLine("Begin to call eBay API, please wait ...")
                apiCall.Execute()
                Console.WriteLine("End to call eBay API, show call result ...")
                Console.WriteLine()

                'Handle the result returned
                Console.WriteLine("UserID: " + apiCall.User.UserID.ToString())

                Console.WriteLine("EIAS Token is: " + apiCall.User.EIASToken.ToString())
                Console.WriteLine()

                If (apiCall.User.eBayGoodStanding = True) Then
                    Console.WriteLine("User has good eBay standing")
                End If

                Console.WriteLine("Rating Star color: " + apiCall.User.FeedbackRatingStar.ToString())
                Console.WriteLine("Feedback score: " + apiCall.User.FeedbackScore.ToString())
                Console.WriteLine()

                Console.WriteLine("Total count of Negative Feedback: " + apiCall.User.UniqueNegativeFeedbackCount.ToString())
                Console.WriteLine("Total count of Neutral Feedback: " + apiCall.User.UniqueNeutralFeedbackCount.ToString())
                Console.WriteLine("Total count of Positive Feedback: " + apiCall.User.UniquePositiveFeedbackCount.ToString())


            Catch ex As Exception
                Console.WriteLine("Failed to get user data : " + ex.Message)
            End Try
        End Sub

        Private Function GetApiContext() As ApiContext
            If (apiContext IsNot Nothing) Then
                Return apiContext
            End If

            Dim apiCredential As ApiCredential = New ApiCredential
            apiCredential.eBayToken = ebayAuthToken

            apiContext = New ApiContext With {
            .SoapApiServerUrl = ApiServerUrl,
            .ApiCredential = apiCredential,
            .Site = SiteCodeType.US,
            .ApiLogManager = New ApiLogManager
        }

            Dim fileLogger As FileLogger = New FileLogger("listing_log.txt", True, True, True)
            apiContext.ApiLogManager.ApiLoggerList.Add(fileLogger)
            apiContext.ApiLogManager.EnableLogging = True

            Return apiContext
        End Function

        Private Function NewAmount(amount As Integer) As AmountType
            Dim amountType As AmountType = New AmountType()
            amountType.Value = amount
            amountType.currencyID = CurrencyCodeType.USD
            Return amountType
        End Function
    End Class
End Namespace