Imports eBay.ApiClient.Auth.OAuth2
Imports eBay.ApiClient.Auth.OAuth2.Model

Public Class EbayAuthentication
    Private configFilePath As String = "ebay-config.yaml"
    Private credentials As CredentialUtil.Credentials

    Public Sub New()
        CredentialUtil.Load(configFilePath)
        credentials = CredentialUtil.GetCredentials(OAuthEnvironment.SANDBOX)
    End Sub
End Class
