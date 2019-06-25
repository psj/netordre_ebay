Imports eBay.Service.Call
Imports eBay.Service.Core.Sdk
Imports eBay.Service.Core.Soap
Imports eBay.Service.Util

Public Class ebayAPI
    Private apiContext As ApiContext = Nothing
    Private ApiServerUrl As String = "https://api.sandbox.ebay.com/wsapi"
    Private ebayAuthToken As String = "AgAAAA**AQAAAA**aAAAAA**Lc4RXQ**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6wFk4aiCZWBpgudj6x9nY+seQ**DAcFAA**AAMAAA**KTUMnytC+yo0/wWROqYZY8Wa8PS3IBBd+QUoL34FbmhJyH5ZqqkXnmeST9YuxYzRsMOGes0brqiKl8bSe7Hf/V8xfOP9+Tgk5UG40wJ74BHrdR06hmlJMivpeDxJ32ggWiOQ0Y3rzEKnnR/b+Z/41K7Rf4zn9BY8VIUlzY5TLXNlcjhyEWSeeRPyM+MpgWZWjPTv0cDvx8PYNmFu7H8eVbDXx0anxtpPh53huUXFpEUsmxTowwOCsYKMj/Xpm9FjTVcr5viBWOSSwrL/JJPOTOgYNf9KxnQOCqRFgUECBlr5JbS5yOGl4AA1QcNx82WYDWeHsH5caL4KxuSfr+jLOV+7hD4RJHLklNFXsHysb1RJAfx25dJ+zbzHcbtf8YEfYxfrI746hwy1LzLw0xL6J/3Li1GiABVcFTSGXP/qvuqiPzG3ZSUid/LNdScOXviIwQy9A2mkdgcEhGKN3NZVS19SO2s3G2QIn2VqrZ41M1OAVfttlgZFhhk9fCHb1Clmi59V18quGyugunSsBQk3Wd/4MT7WdSo4qWHgT+3vNipWkYIRBe6V8vHR+p0Zk/W1ujzpToPUrWfihw+qzwEDnabkZxsHw31qF3D/agU81rOQNI1sMm5jyM99Fe4x8bzkqZzafkhfvhj38FzzszM2vInC+eGmUFnZYxVd7aAmoIhab2fyWxtuXVloxIkF0jcInc4Ytxysn6Z/31Ln1MCQpAhEOTE1yysDEryvYkNlP4tLfnNFI9MN/kCYxu9Lv8/t"

    Public Sub New()
        GetApiContext()
    End Sub

    Public Sub ListItem()
        Try
            Dim item As ItemType = BuildItem()

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
            Console.WriteLine(String.Format("Listing fee is: {0}", listingFee))
            Console.WriteLine(String.Format("Listed Item ID: {0}", item.ItemID))
        Catch ex As Exception
            Console.WriteLine("Failed to get user data : " + ex.Message)
        End Try

    End Sub

    Private Function BuildItem() As ItemType
        Dim item As ItemType = New ItemType()

        item.Title = "Test Item"
        item.Description = "eBay SDK sample test item"

        item.ListingType = ListingTypeCodeType.FixedPriceItem
        item.Currency = CurrencyCodeType.GBP
        item.ListingDuration = "Days_3"
        item.Location = "Hedehusene"
        item.Country = CountryCodeType.DK

        Dim category As CategoryType = New CategoryType()
        category.CategoryID = "4174"
        item.PrimaryCategory = category

        item.PaymentMethods = New BuyerPaymentMethodCodeTypeCollection()
        item.PaymentMethods.Add(BuyerPaymentMethodCodeType.PayPal)
        item.PayPalEmailAddress = "me@ebay.com"

        item.DispatchTimeMax = 1
        item.ShippingDetails = BuildShippingDetails()

        item.ReturnPolicy = New ReturnPolicyType()
        item.ReturnPolicy.ReturnsAcceptedOption = "ReturnsAccepted"

        Return item
    End Function

    Private Function BuildShippingDetails() As ShippingDetailsType
        Dim shipping_details As ShippingDetailsType = New ShippingDetailsType()

        shipping_details.ApplyShippingDiscount = True

        Dim amount As AmountType = New AmountType()
        amount.Value = 30
        amount.currencyID = CurrencyCodeType.GBP
        shipping_details.PaymentInstructions = "eBay .Net SDK test instruction."

        shipping_details.ShippingType = ShippingTypeCodeType.Flat
        shipping_details.ShippingServiceUsed = ShippingServiceCodeType.UK_RoyalMailStandardParcel

        'Dim shippingOptions As ShippingServiceOptionsType = New ShippingServiceOptionsType()
        'shippingOptions.ShippingService = ShippingServiceCodeType.UK_RoyalMailStandardParcel

        'amount = New AmountType()
        'amount.Value = 2
        'amount.currencyID = CurrencyCodeType.GBP
        'shippingOptions.ShippingServiceAdditionalCost = amount

        'amount = New AmountType()
        'amount.Value = 100
        'amount.currencyID = CurrencyCodeType.GBP
        'shippingOptions.ShippingServiceCost = amount
        'shippingOptions.ShippingServicePriority = 1

        'amount = New AmountType()
        'amount.Value = 25
        'amount.currencyID = CurrencyCodeType.GBP
        'shippingOptions.ShippingInsuranceCost = amount

        'shipping_details.ShippingServiceOptions = New ShippingServiceOptionsTypeCollection()
        'shipping_details.ShippingServiceOptions.Add(shippingOptions)

        Return shipping_details

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
        If (ApiContext IsNot Nothing) Then
            Return apiContext
        End If

        Dim apiCredential As ApiCredential = New ApiCredential
        apiCredential.eBayToken = ebayAuthToken

        apiContext = New ApiContext With {
            .SoapApiServerUrl = ApiServerUrl,
            .ApiCredential = apiCredential,
            .Site = SiteCodeType.UK,
            .ApiLogManager = New ApiLogManager
        }

        Dim fileLogger As FileLogger = New FileLogger("listing_log.txt", True, True, True)
        apiContext.ApiLogManager.ApiLoggerList.Add(fileLogger)
        apiContext.ApiLogManager.EnableLogging = True

        Return apiContext
    End Function

End Class
