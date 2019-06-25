Imports eBay.ApiClient.Auth.OAuth2
Imports eBay.ApiClient.Auth.OAuth2.Model

Public Class EbayAuthentication
    Private configFilePath As String = "ebay-config.yaml"
    Private credentials As CredentialUtil.Credentials
    Private scopes As New List(Of String)(
        {
            "https://api.ebay.com/oauth/api_scope"
        }
    )

    Public Sub New()
        CredentialUtil.Load(configFilePath)
        credentials = CredentialUtil.GetCredentials(OAuthEnvironment.SANDBOX)
    End Sub

    Public Function FetchApplicationToken() As String
        Dim response As OAuthResponse
        Dim oAuth2Api As OAuth2Api = New OAuth2Api

        response = oAuth2Api.GetApplicationToken(OAuthEnvironment.SANDBOX, scopes)
        Return response.AccessToken.ToString
    End Function
End Class
