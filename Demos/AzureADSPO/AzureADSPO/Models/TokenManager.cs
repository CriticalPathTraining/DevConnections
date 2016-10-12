using System;
using System.Configuration;
using System.Security.Claims;
using System.Web;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;

namespace AzureADSPO.Models {
  public class TokenManager {
    public static async Task<string> GetAccessToken(string resource) {

      // get ClaimsPrincipal for current user
      ClaimsPrincipal currentUserClaims = ClaimsPrincipal.Current;
      string signedInUserID = currentUserClaims.FindFirst(ClaimTypes.NameIdentifier).Value;
      string tenantID = currentUserClaims.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
      string userObjectID = currentUserClaims.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

      ApplicationDbContext db = new ApplicationDbContext();
      ADALTokenCache userTokenCache = new ADALTokenCache(signedInUserID);

      string urlAuthorityRoot = ConfigurationManager.AppSettings["ida:AADInstance"];
      string urlAuthorityTenant = urlAuthorityRoot + tenantID;

      AuthenticationContext authenticationContext =
        new AuthenticationContext(urlAuthorityTenant, userTokenCache);

      Uri uriReplyUrl = new Uri(HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Path));

      string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
      string clientSecret = ConfigurationManager.AppSettings["ida:ClientSecret"];
      ClientCredential clientCredential = new ClientCredential(clientId, clientSecret);

      UserIdentifier userIdentifier = new UserIdentifier(userObjectID, UserIdentifierType.UniqueId);

      AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenSilentAsync(resource, clientCredential, userIdentifier);

      return authenticationResult.AccessToken;

    }

  }
}