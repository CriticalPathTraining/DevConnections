using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Web;

namespace AzureGraphAdalWebClient {

  public class DemoConstants {

    public const string ClientId = "3c8db355-8f09-4077-bdd2-e3472164cf65";
    public const string ClientSecret = "Fhh3gc47jlONJ32rsq43O7EWL7cHYjXaWT54xPA6p9Y=";
    public static readonly string ClientSecretEncoded = HttpUtility.UrlEncode(ClientSecret);

    public const string TargetResource = "https://graph.windows.net";

    public const string ClientUrl = "https://localhost:44300/";
    public const string ClientReplyUrl = "https://localhost:44300/ReplyUrl/";

    public const string urlAuthorizationEndpoint = "https://login.windows.net/common/oauth2/authorize" +
                                                      "?resource=" + TargetResource +
                                                      "&client_id=" + ClientId +
                                                      "&redirect_uri=" + ClientReplyUrl +
                                                      "&prompt=consent" + 
                                                      "&response_type=code";

    public const string urlTokenEndpoint = "https://login.windows.net/common/oauth2/token";

    public static string AccessTokenRequestBody = "grant_type=authorization_code" +
                                                   "&resource=" + TargetResource +
                                                   "&redirect_uri=" + ClientReplyUrl +
                                                   "&client_id=" + ClientId +
                                                   "&client_secret=" + ClientSecretEncoded +
                                                   "&code=";

    public static ClientAssertionCertificate GetClientAssertionCertificate() {
      string PfxFilePassword = "pass@word1";
      string PfxFilePath = HttpContext.Current.Server.MapPath("~/Secrets/AppOnlyClientCertificate01.pfx");
      X509Certificate2 privateKeyCert = new X509Certificate2(PfxFilePath, PfxFilePassword);
      return new ClientAssertionCertificate(ClientId, privateKeyCert);
    }


  }
}