using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using AcsSPOWeb.Models;

namespace AcsSPOWeb.Controllers
{
    public class SharePointSiteController : Controller
    {
    public ActionResult Index() {

      SharePointSite site = SharePointSiteManager.GetSharePointSite();
      return View(site);
    }

    public ActionResult Lists() {

      IEnumerable<SharePointList> lists = SharePointSiteManager.GetLists();

      return View(lists);
    }

    public ActionResult CreateCustomersList() {

      SharePointList list = SharePointSiteManager.CreateCustomersList();

      return View(list);
    }
  }
}