using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;

using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace AzureADSPO.Models {

  public class SharePointSite {
    public string Id { get; set; }
    public string Title { get; set; }
    public string Url { get; set; }
    public string ServerRelativeUrl { get; set; }
    public string WebTemplate { get; set; }
    public int Configuration { get; set; }
    public string MasterUrl { get; set; }
    public string CustomMasterUrl { get; set; }
    public string AlternateCssUrl { get; set; }
    public bool EnableMinimalDownload { get; set; }
    public int Language { get; set; }
    public string LastItemModifiedDate { get; set; }
  }

  public class SharePointList {
    public string Id { get; set; }
    public string Title { get; set; }
    public string DefaultViewUrl { get; set; }
  }
  
  public class SharePointSiteManager {

    static readonly string siteUrl = ConfigurationManager.AppSettings["targetSPOSite"];

    private static async Task<ClientContext> GetClientContext() {
      ClientContext clientContext = new ClientContext(siteUrl);
      // connect access token to all outbound requests
      string accessToken = await TokenManager.GetAccessToken(siteUrl);
      clientContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e) {
        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
      };
      return clientContext;
    }

    public static async Task<SharePointSite> GetSharePointSite() {

      ClientContext ctx = await GetClientContext();
      ctx.Load(ctx.Web);
      ctx.ExecuteQuery();

      SharePointSite site = new SharePointSite {
        Id = ctx.Web.Id.ToString(),
        Title = ctx.Web.Title,
        Url = ctx.Web.Url,
        ServerRelativeUrl = ctx.Web.ServerRelativeUrl,
        WebTemplate = ctx.Web.WebTemplate,
        Configuration = ctx.Web.Configuration,
        MasterUrl = ctx.Web.MasterUrl,
        CustomMasterUrl = ctx.Web.CustomMasterUrl,
        AlternateCssUrl = ctx.Web.AlternateCssUrl,
        EnableMinimalDownload = ctx.Web.EnableMinimalDownload,
        Language = (int)ctx.Web.Language,
        LastItemModifiedDate = ctx.Web.LastItemModifiedDate.ToString()
      };

      ctx.Dispose();
        
      return site;
 
    }

    public static async Task<IEnumerable<SharePointList>> GetLists() {
      ClientContext ctx = await GetClientContext();
      ListCollection Lists = ctx.Web.Lists;
      ctx.Load(Lists, siteLists => siteLists.Where(list => !list.Hidden)
                                            .Include(list => list.Id, list => list.Title, list => list.DefaultViewUrl));
      ctx.ExecuteQuery();

      List<SharePointList> lists = new List<SharePointList>();
      foreach (var list in ctx.Web.Lists) {
        lists.Add(new SharePointList {
          Id = list.Id.ToString(),
          Title = list.Title,
          DefaultViewUrl = siteUrl + "/" + list.DefaultViewUrl
        });
      }
      
      ctx.Dispose();
      return lists;
    }

    public static async Task<SharePointList> CreateCustomersList() {
      ClientContext ctx = await GetClientContext();
      ctx.Load(ctx.Web);
      ctx.ExecuteQuery();
      string listTitle = "Customers";

      // delete list if it exists
      ExceptionHandlingScope scope = new ExceptionHandlingScope(ctx);
      using (scope.StartScope()) {
        using (scope.StartTry()) {
          ctx.Web.Lists.GetByTitle(listTitle).DeleteObject();
        }
        using (scope.StartCatch()) { }
      }

      // create and initialize ListCreationInformation object
      ListCreationInformation listInformation = new ListCreationInformation();
      listInformation.Title = listTitle;
      listInformation.Url = "Lists/Customers";
      listInformation.QuickLaunchOption = QuickLaunchOptions.On;
      listInformation.TemplateType = (int)ListTemplateType.Contacts;

      // Add ListCreationInformation to lists collection and return list object
      List list = ctx.Web.Lists.Add(listInformation);

      // modify additional list properties and update
      list.OnQuickLaunch = true;
      list.EnableAttachments = false;
      list.Update();

      // send command to server to create list
      ctx.Load(list, l => l.Id, l => l.Title, l => l.DefaultViewUrl);
      ctx.ExecuteQuery();

      // add an item to the list
      ListItemCreationInformation lici1 = new ListItemCreationInformation();
      var item1 = list.AddItem(lici1);
      item1["Title"] = "Lennon";
      item1["FirstName"] = "John";
      item1.Update();

      // add a second item 
      ListItemCreationInformation lici2 = new ListItemCreationInformation();
      var item2 = list.AddItem(lici2);
      item2["Title"] = "McCartney";
      item2["FirstName"] = "Paul";
      item2.Update();

      // send add commands to server
      ctx.ExecuteQuery();

      SharePointList newList = new SharePointList {
        Id = list.Id.ToString(),
        Title = list.Title,
        DefaultViewUrl = siteUrl + "/" + list.DefaultViewUrl
      };

      ctx.Dispose();

      return newList;

    }
  }
}