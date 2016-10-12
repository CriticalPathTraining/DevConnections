using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using AcsSPOWeb.Models;

namespace AcsSPOWeb.Controllers {
  public class HomeController : Controller {

    public ActionResult Index() {
      ViewBag.Message = "Hello MVC.";
      return View();
    }

    public ActionResult SharePointSession() {
      return View(SharePointSessionManager.GetSharePointSessionState());
    }

  }
}
