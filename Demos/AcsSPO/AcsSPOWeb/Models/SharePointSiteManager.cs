using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace AcsSPOWeb.Models {

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

    static readonly string siteUrl = "https://cbd365labs.sharepoint.com/sites/dev";

    private static ClientContext GetClientContext() {
      ClientContext clientContext = new ClientContext(siteUrl);
      // connect access token to all outbound requests
      string accessToken = SharePointSessionManager.GetSharePointSessionState().AccessToken;
      clientContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e) {
        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
      };
      return clientContext;
    }

    public static SharePointSite GetSharePointSite() {

      ClientContext ctx = GetClientContext();
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

    public static IEnumerable<SharePointList> GetLists() {
      ClientContext ctx = GetClientContext();
      ListCollection Lists = ctx.Web.Lists;
      ctx.Load(Lists, siteLists => siteLists.Where(list => !list.Hidden)
                                            .Include(list => list.Id, list => list.Title, list => list.DefaultViewUrl));
      ctx.ExecuteQuery();


      string urlAuthority = "https://" + (new Uri(siteUrl)).Authority;

      List<SharePointList> lists = new List<SharePointList>();
      foreach (var list in ctx.Web.Lists) {
        lists.Add(new SharePointList {
          Id = list.Id.ToString(),
          Title = list.Title,
          DefaultViewUrl = urlAuthority + list.DefaultViewUrl
        });
      }

      ctx.Dispose();
      return lists;
    }

    public static SharePointList CreateCustomersList() {
      ClientContext ctx = GetClientContext();
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
      ctx.Load(list, l => l.Id , l => l.Title, l => l.DefaultViewUrl);
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

      string urlAuthority = "https://" + (new Uri(siteUrl)).Authority;

      SharePointList newList = new SharePointList {
        Id = list.Id.ToString(),
        Title = list.Title,
        DefaultViewUrl = urlAuthority + list.DefaultViewUrl
      };

      ctx.Dispose();

      return newList;

    }
  }
}