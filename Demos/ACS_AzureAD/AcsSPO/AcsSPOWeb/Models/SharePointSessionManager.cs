using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;
using System.Web.SessionState;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Newtonsoft.Json;

namespace AcsSPOWeb.Models {

  public class SharePointSessionState {
    public string RemoteWebUrl { get; set; }
    public string HostWebUrl { get; set; }
    public string HostWebDomain { get; set; }
    public string HostWebTitle { get; set; }
    public string HostTenantId { get; set; }
    public string CurrentUserName { get; set; }
    public string CurrentUserEmail { get; set; }
    public string TargetResource { get; set; }
    public string RefreshToken { get; set; }
    public DateTime RefreshTokenExpires { get; set; }
    public string AccessToken { get; set; }
    public DateTime AccessTokenExpires { get; set; }
  }

  public class SharePointUserResult {
    public string Title { get; set; }
    public string Email { get; set; }
    public string IsSiteAdmin { get; set; }
  }

  public class SharePointSessionManager {

    static HttpRequest request = HttpContext.Current.Request;
    static HttpSessionState session = HttpContext.Current.Session;
    static SharePointSessionState sessionState = new SharePointSessionState();

    private static string ExecuteGetRequest(string restUri, string accessToken) {

      // setup request 
      HttpClient client = new HttpClient();
      client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
      client.DefaultRequestHeaders.Add("Accept", "application/json");
      // execute request 
      HttpResponseMessage response = client.GetAsync(restUri).Result;
      // handle response 
      if (response.IsSuccessStatusCode) {
        return response.Content.ReadAsStringAsync().Result;
      }
      else {
        // ERROR during HTTP GET operation 
        return string.Empty;
      }
    }

    private static void AuthenticateUser() {

      sessionState.RemoteWebUrl = request.Url.Authority;
      sessionState.HostWebUrl = request["SPHostUrl"];
      sessionState.HostWebDomain = (new Uri(sessionState.HostWebUrl)).Authority;
      sessionState.HostWebTitle = request.Form["SPSiteTitle"];
      string contextTokenString = request.Form["SPAppToken"];

      // create SharePoint context token object 
      SharePointContextToken contextToken =
        TokenHelper.ReadAndValidateContextToken(contextTokenString, sessionState.RemoteWebUrl);
      // read session state from SharePoint context token object 
      sessionState.HostTenantId = contextToken.Realm;
      sessionState.TargetResource = contextToken.Audience;
      sessionState.RefreshToken = contextToken.RefreshToken;
      sessionState.RefreshTokenExpires = contextToken.ValidTo;

      // use refresh token to acquire access token response from Azure ACS 
      OAuth2AccessTokenResponse AccessTokenResponse =
        TokenHelper.GetAccessToken(contextToken, sessionState.HostWebDomain);
      // Read access token and ExpiresOn value from access token response 
      sessionState.AccessToken = AccessTokenResponse.AccessToken;
      sessionState.AccessTokenExpires = AccessTokenResponse.ExpiresOn;

      // call SharePoint REST API to get information about current user 
      string restUri = sessionState.HostWebUrl + "/_api/web/currentUser/";
      string jsonCurrentUser = ExecuteGetRequest(restUri, sessionState.AccessToken);

      // convert json result to strongly-typed C# object 
      SharePointUserResult userResult = JsonConvert.DeserializeObject<SharePointUserResult>(jsonCurrentUser);

      sessionState.CurrentUserName = userResult.Title;
      sessionState.CurrentUserEmail = userResult.Email;

      // write session state out to ASP.NET session object 
      session["SharePointSessionState"] = sessionState;

      // update UserIsAuthenticated session variable 
      session["UserIsAuthenticated"] = "true";
    }

    private static bool UserIsAuthentiated() {
      return (session["UserIsAuthenticated"] != null) &&
             (session["UserIsAuthenticated"].Equals("true"));
    }

    public static void InitializeRequest(ControllerBase controller) {

      if (!UserIsAuthentiated()) {
        AuthenticateUser();
      }
      else {
        // if user is authenticated, copy session state from previous request 
        sessionState = (SharePointSessionState)session["SharePointSessionState"];
      }

      // add session state to ViewBag to make it accessible in views 
      controller.ViewBag.HostWebUrl = sessionState.HostWebUrl;
      controller.ViewBag.HostWebTitle = sessionState.HostWebTitle;
      controller.ViewBag.CurrentUserName = sessionState.CurrentUserName;
    }

    public static SharePointSessionState GetSharePointSessionState() {
      return sessionState;
    }

  }

}