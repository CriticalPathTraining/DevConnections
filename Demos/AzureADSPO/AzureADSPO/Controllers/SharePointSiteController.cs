using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using AzureADSPO.Models;
using System.Threading.Tasks;

namespace AzureADSPO.Controllers {

  [Authorize]
  public class SharePointSiteController : Controller {

    // GET: SharePointSite
    public async Task<ActionResult> Index() {

      SharePointSite site = await SharePointSiteManager.GetSharePointSite();

      return View(site);
    }

    public async Task<ActionResult> Lists() {

      IEnumerable<SharePointList> lists = await SharePointSiteManager.GetLists();

      return View(lists);
    }

    public async Task<ActionResult> CreateCustomersList() {

      SharePointList list = await SharePointSiteManager.CreateCustomersList();

      return View(list);
    }
  }
}